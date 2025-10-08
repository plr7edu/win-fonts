import os
from pathlib import Path
from fontTools.ttLib import TTFont, TTCollection
import sys
from datetime import datetime
from collections import defaultdict

class FontAssemblerError(Exception):
    """Custom exception for font assembler errors."""
    pass

def get_font_info(font_path):
    """Extract detailed font information from a TTF file."""
    try:
        font = TTFont(font_path)
        name_table = font['name']
        
        family_name = None
        subfamily_name = None
        full_name = None
        
        # Extract names with priority for Windows platform
        for record in name_table.names:
            text = record.toUnicode()
            
            # Preferred Family (ID 16) - highest priority
            if record.nameID == 16 and not family_name:
                family_name = text
            # Font Family (ID 1) - fallback
            elif record.nameID == 1 and not family_name:
                family_name = text
            # Subfamily (ID 2) - style like Bold, Italic
            elif record.nameID == 2 and not subfamily_name:
                subfamily_name = text
            # Full name (ID 4)
            elif record.nameID == 4 and not full_name:
                full_name = text
        
        # Get font metrics for validation
        num_glyphs = font['maxp'].numGlyphs if 'maxp' in font else 0
        
        font.close()
        
        return {
            'family': family_name,
            'subfamily': subfamily_name or 'Regular',
            'full_name': full_name or family_name,
            'num_glyphs': num_glyphs
        }
    except Exception as e:
        raise FontAssemblerError(f"Error reading font '{font_path.name}': {e}")

def sanitize_family_name(family_name):
    """Remove spaces and invalid characters from family name."""
    # Remove spaces
    sanitized = family_name.replace(' ', '')
    # Remove invalid filename characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        sanitized = sanitized.replace(char, '')
    return sanitized.strip()

def sanitize_style_name(style_name):
    """Normalize style name."""
    # Remove spaces and standardize
    sanitized = style_name.replace(' ', '')
    return sanitized.strip()

def rename_font_files(ttf_files):
    """Rename all font files to <FontFamily>-<Weight>-<Style>.ttf pattern."""
    print(f"\n{'='*60}")
    print(f"STEP 1: RENAMING FONT FILES")
    print(f"{'='*60}\n")
    
    renamed_files = []
    rename_map = {}
    
    for ttf_file in ttf_files:
        try:
            # Get font info
            info = get_font_info(ttf_file)
            
            # Sanitize names
            family = sanitize_family_name(info['family'])
            style = sanitize_style_name(info['subfamily'])
            
            # Create new filename
            new_filename = f"{family}-{style}.ttf"
            new_path = ttf_file.parent / new_filename
            
            # Check if rename is needed
            if ttf_file.name != new_filename:
                # Check if target already exists
                if new_path.exists() and new_path != ttf_file:
                    print(f"⚠ WARNING: Target file already exists: {new_filename}")
                    print(f"  Skipping rename for: {ttf_file.name}")
                    renamed_files.append(ttf_file)
                    rename_map[ttf_file] = ttf_file
                else:
                    # Rename the file
                    ttf_file.rename(new_path)
                    print(f"✓ Renamed: {ttf_file.name}")
                    print(f"       to: {new_filename}")
                    renamed_files.append(new_path)
                    rename_map[ttf_file] = new_path
            else:
                print(f"✓ Already correct: {ttf_file.name}")
                renamed_files.append(ttf_file)
                rename_map[ttf_file] = ttf_file
                
        except Exception as e:
            print(f"✗ ERROR renaming {ttf_file.name}: {e}")
            # Keep original file in list
            renamed_files.append(ttf_file)
            rename_map[ttf_file] = ttf_file
    
    print(f"\n✓ Renamed {len([k for k, v in rename_map.items() if k != v])} file(s)")
    return renamed_files

def group_fonts_by_family(ttf_files):
    """Group font files by their font family."""
    print(f"\n{'='*60}")
    print(f"STEP 2: GROUPING FONTS BY FAMILY")
    print(f"{'='*60}\n")
    
    family_groups = defaultdict(list)
    
    for ttf_file in ttf_files:
        try:
            info = get_font_info(ttf_file)
            info['path'] = ttf_file
            
            # Use sanitized family name as key
            family_key = sanitize_family_name(info['family'])
            family_groups[family_key].append(info)
            
        except Exception as e:
            print(f"✗ ERROR reading {ttf_file.name}: {e}")
            continue
    
    # Display grouped fonts
    print(f"Found {len(family_groups)} font familie(s):\n")
    
    for family_name, fonts in sorted(family_groups.items()):
        print(f"📁 Family: {family_name} ({len(fonts)} variant(s))")
        for font_info in fonts:
            print(f"   ├─ {font_info['path'].name}")
            print(f"   │  Style: {font_info['subfamily']} | Glyphs: {font_info['num_glyphs']}")
        print()
    
    return family_groups

def validate_font_group(family_name, fonts):
    """Validate fonts in a group."""
    # Check for duplicate styles
    styles = [f['subfamily'] for f in fonts]
    duplicates = set([s for s in styles if styles.count(s) > 1])
    
    if duplicates:
        print(f"⚠ WARNING [{family_name}]: Duplicate styles detected: {', '.join(duplicates)}")
        print(f"  The collection will include all files, but this may cause conflicts.\n")
    
    return True

def create_font_collection(family_name, font_info_list, output_path):
    """Create TrueType Collection file for a font family."""
    try:
        print(f"\n{'─'*60}")
        print(f"Creating: {family_name}.ttc")
        print(f"{'─'*60}")
        
        # Load all fonts
        fonts = []
        for info in font_info_list:
            print(f"  Loading: {info['path'].name}")
            fonts.append(TTFont(str(info['path'])))
        
        # Create collection
        collection = TTCollection()
        collection.fonts = fonts
        
        # Save
        collection.save(str(output_path))
        
        # Close all fonts
        for font in fonts:
            font.close()
        
        # Get file size
        file_size = output_path.stat().st_size / 1024  # KB
        
        print(f"  ✓ Saved: {output_path.name}")
        print(f"  ✓ Size: {file_size:.2f} KB")
        print(f"  ✓ Variants: {len(fonts)}")
        
        return True
        
    except Exception as e:
        print(f"  ✗ ERROR: Failed to create collection: {e}")
        return False

def main():
    """Main function to orchestrate the font assembly process."""
    print(f"\n{'='*60}")
    print(f"MULTI-FAMILY FONTS ASSEMBLER")
    print(f"{'='*60}")
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    try:
        # Define paths
        current_path = Path.cwd()
        source_folder = current_path / "1_source"
        export_folder = current_path / "2_export"
        
        # Validate folders
        if not source_folder.exists():
            raise FontAssemblerError(f"Source folder '{source_folder}' does not exist!")
        
        # Create export folder if it doesn't exist
        export_folder.mkdir(exist_ok=True)
        print(f"Source: {source_folder}")
        print(f"Export: {export_folder}")
        
        # Get all files
        all_files = list(source_folder.glob("*"))
        
        if not all_files:
            raise FontAssemblerError(f"No files found in '{source_folder}'!")
        
        # Filter TTF files (case insensitive)
        ttf_files = [f for f in all_files if f.suffix.lower() == '.ttf' and f.is_file()]
        
        if not ttf_files:
            raise FontAssemblerError("No .ttf files found in source folder!")
        
        # Check for non-TTF font files
        font_extensions = {'.otf', '.woff', '.woff2', '.ttc', '.eot'}
        other_fonts = [f for f in all_files if f.suffix.lower() in font_extensions]
        
        if other_fonts:
            print(f"\n⚠ WARNING: Non-TTF font files found (will be ignored):")
            for f in other_fonts:
                print(f"  - {f.name} ({f.suffix.upper()})")
        
        print(f"\nFound {len(ttf_files)} TTF file(s) to process")
        
        # STEP 1: Rename files to standard pattern
        renamed_files = rename_font_files(ttf_files)
        
        # STEP 2: Group fonts by family
        family_groups = group_fonts_by_family(renamed_files)
        
        if not family_groups:
            raise FontAssemblerError("No valid font families found after grouping!")
        
        # STEP 3: Create TTC files for each family
        print(f"{'='*60}")
        print(f"STEP 3: CREATING FONT COLLECTIONS")
        print(f"{'='*60}")
        
        successful = 0
        failed = 0
        
        for family_name, fonts in sorted(family_groups.items()):
            # Validate group
            validate_font_group(family_name, fonts)
            
            # Create output path
            output_filename = f"{family_name}.ttc"
            output_path = export_folder / output_filename
            
            # Check if output already exists
            if output_path.exists():
                print(f"\n⚠ WARNING: Output file already exists and will be overwritten:")
                print(f"  {output_path.name}")
            
            # Create collection
            if create_font_collection(family_name, fonts, output_path):
                successful += 1
            else:
                failed += 1
        
        # Final summary
        print(f"\n{'='*60}")
        print(f"✓ PROCESS COMPLETED")
        print(f"{'='*60}")
        print(f"Total families processed: {len(family_groups)}")
        print(f"Successful: {successful}")
        print(f"Failed: {failed}")
        print(f"Output location: {export_folder}")
        print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*60}\n")
        
        if failed > 0:
            sys.exit(1)
        
    except FontAssemblerError as e:
        print(f"\n{'='*60}")
        print(f"ERROR")
        print(f"{'='*60}")
        print(f"{e}")
        print(f"{'='*60}\n")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n⚠ Operation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n{'='*60}")
        print(f"UNEXPECTED ERROR")
        print(f"{'='*60}")
        print(f"{e}")
        import traceback
        traceback.print_exc()
        print(f"{'='*60}\n")
        sys.exit(1)

if __name__ == "__main__":
    main()

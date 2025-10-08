import os
import sys
from pathlib import Path
from fontTools import ttLib
from fontTools.ttLib import TTFont, TTCollection
from fontTools.varLib import instancer

def is_collection_font(font_path):
    """Check if a font file is a collection (has multiple fonts inside)."""
    try:
        # Check if it's a TTC file
        if font_path.suffix.lower() == '.ttc':
            return True
        
        # Check if it's a variable font with instances
        font = TTFont(str(font_path))
        has_instances = 'fvar' in font and len(font['fvar'].instances) > 0
        font.close()
        
        if has_instances:
            return True
        
        # Try to open as collection
        try:
            ttc = TTCollection(str(font_path))
            if len(ttc.fonts) > 1:
                ttc.close()
                return True
            ttc.close()
        except:
            pass
        
        return False
    except Exception as e:
        return False

def separate_font_collection(font_path, output_dir):
    """Separate a TrueType Collection (.ttc) into individual .ttf files."""
    try:
        ttc = TTCollection(str(font_path))
        count = 0
        
        for i, font in enumerate(ttc.fonts):
            # Try to get the font subfamily name (Bold, Regular, etc.)
            if 'name' in font:
                name_record = font['name']
                subfamily = name_record.getDebugName(2) or f"Variant{i+1}"
                family = name_record.getDebugName(1) or "Font"
            else:
                family = "Font"
                subfamily = f"Variant{i+1}"
            
            # Clean the name for filename
            safe_name = f"{family}-{subfamily}".replace(" ", "").replace("/", "-")
            output_file = output_dir / f"{safe_name}.ttf"
            
            # Avoid overwriting files with same name
            counter = 1
            while output_file.exists():
                output_file = output_dir / f"{safe_name}_{counter}.ttf"
                counter += 1
            
            # Save the individual font
            font.save(str(output_file))
            print(f"      [OK] Extracted: {output_file.name}")
            count += 1
        
        ttc.close()
        return count
    except Exception as e:
        print(f"      [ERROR] Failed to process font collection: {e}")
        return 0

def separate_variable_font(font_path, output_dir):
    """Separate a variable font with named instances into individual static .ttf files."""
    try:
        font = TTFont(str(font_path))
        
        # Check if it's a variable font with fvar table
        if 'fvar' not in font:
            font.close()
            return 0
        
        fvar = font['fvar']
        instances = fvar.instances
        
        if len(instances) == 0:
            font.close()
            return 0
        
        count = 0
        
        # Get family name
        if 'name' in font:
            name_record = font['name']
            family = name_record.getDebugName(1) or "Font"
        else:
            family = "Font"
        
        font.close()
        
        for i, instance in enumerate(instances):
            # Reload font for each instance
            var_font = TTFont(str(font_path))
            
            # Get instance subfamily name
            subfamily = var_font['name'].getDebugName(instance.subfamilyNameID)
            if not subfamily:
                subfamily = f"Instance{i+1}"
            
            # Clean the name for filename
            safe_name = f"{family}-{subfamily}".replace(" ", "").replace("/", "-")
            output_file = output_dir / f"{safe_name}.ttf"
            
            # Avoid overwriting files with same name
            counter = 1
            while output_file.exists():
                output_file = output_dir / f"{safe_name}_{counter}.ttf"
                counter += 1
            
            try:
                # Create location dict for this instance
                location = {axis.axisTag: instance.coordinates[axis.axisTag] 
                           for axis in var_font['fvar'].axes}
                
                # Instantiate the variable font to this location
                static_font = instancer.instantiateVariableFont(var_font, location)
                
                # Save the static instance
                static_font.save(str(output_file))
                print(f"      [OK] Extracted: {output_file.name}")
                count += 1
                static_font.close()
            except Exception as e:
                print(f"      [ERROR] Failed to instantiate {subfamily}: {e}")
                var_font.close()
                continue
        
        return count
    except Exception as e:
        print(f"      [ERROR] Failed to process variable font: {e}")
        return 0

def extract_font(font_path, output_dir, indent_level=2):
    """Extract a single font file and return number of variants extracted."""
    extracted_count = 0
    
    # Check if it's a TTC (collection) or regular TTF
    if font_path.suffix.lower() == '.ttc':
        extracted_count = separate_font_collection(font_path, output_dir)
    else:
        # Try to separate as variable font first
        extracted_count = separate_variable_font(font_path, output_dir)
        
        # If not a variable font, try as collection
        if extracted_count == 0:
            try:
                extracted_count = separate_font_collection(font_path, output_dir)
            except:
                pass
    
    return extracted_count

def process_single_font(font_path, parent_dir, level=0):
    """
    Process a single font file: extract it and recursively process results.
    This ensures depth-first processing - fully complete one font before moving to next.
    """
    indent = "  " * level
    total_extracted = 0
    
    print(f"{indent}[PROCESS] {font_path.name}")
    
    # Create folder with same name as font file (without extension)
    folder_name = font_path.stem
    output_dir = parent_dir / folder_name
    output_dir.mkdir(exist_ok=True)
    print(f"{indent}  [INFO] Created folder: {folder_name}/")
    
    # Extract the font
    extracted_count = extract_font(font_path, output_dir, level + 2)
    
    if extracted_count == 0:
        print(f"{indent}  [WARN] No variants extracted (single-weight font)")
        try:
            output_dir.rmdir()
        except:
            pass
        return 0
    
    print(f"{indent}  [OK] Extracted {extracted_count} variant(s)")
    total_extracted += extracted_count
    
    # Process each extracted font in this folder depth-first
    print(f"{indent}  [INFO] Checking extracted fonts for nested collections...")
    
    # Get all font files in the newly created folder
    extracted_fonts = sorted(list(output_dir.glob("*.ttf")) + 
                           list(output_dir.glob("*.ttc")) + 
                           list(output_dir.glob("*.otf")))
    
    # Check each extracted font for collections
    nested_total = 0
    for extracted_font in extracted_fonts:
        if is_collection_font(extracted_font):
            print(f"{indent}  [FOUND] Nested collection: {extracted_font.name}")
            # Recursively process this font completely before moving to next
            nested_count = process_single_font(extracted_font, output_dir, level + 2)
            nested_total += nested_count
    
    if nested_total > 0:
        print(f"{indent}  [OK] Extracted {nested_total} additional nested variant(s)")
        total_extracted += nested_total
    else:
        print(f"{indent}  [OK] No nested collections found")
    
    return total_extracted

def process_directory(directory, level=0):
    """Process all collection fonts in a directory (non-recursive at this level)."""
    indent = "  " * level
    total_extracted = 0
    
    # Find all font files in this directory only (not subdirectories)
    font_files = (list(directory.glob("*.ttf")) + 
                 list(directory.glob("*.ttc")) + 
                 list(directory.glob("*.otf")))
    
    if len(font_files) == 0:
        return 0
    
    # Identify collection fonts
    collection_fonts = [f for f in font_files if is_collection_font(f)]
    
    if len(collection_fonts) == 0:
        if level == 0:
            print(f"{indent}[INFO] No collection fonts found in this directory")
        return 0
    
    print(f"{indent}[INFO] Found {len(collection_fonts)} collection font(s) to extract")
    print()
    
    # Process each collection font depth-first
    for font_path in sorted(collection_fonts):
        extracted = process_single_font(font_path, directory, level)
        total_extracted += extracted
        print()
    
    return total_extracted

def main():
    print("=" * 80)
    print("Font Separator - Depth-First Recursive Processing")
    print("=" * 80)
    print()
    
    # Check for !source+output directory
    current_dir = Path.cwd()
    source_output_dir = current_dir / "!source+output"
    
    if not source_output_dir.exists():
        print("ERROR: '!source+output' folder not found in the current directory.")
        print(f"Current directory: {current_dir}")
        input("\nPress Enter to exit...")
        sys.exit(1)
    
    print(f"[INFO] Source directory: {source_output_dir}")
    print()
    print("[START] Beginning depth-first extraction process...")
    print("=" * 80)
    print()
    
    # Start processing
    total_extracted = process_directory(source_output_dir)
    
    # Final verification
    print("=" * 80)
    print("[VERIFY] Performing final verification...")
    print("=" * 80)
    
    all_fonts = (list(source_output_dir.rglob("*.ttf")) + 
                list(source_output_dir.rglob("*.ttc")) + 
                list(source_output_dir.rglob("*.otf")))
    remaining_collections = [f for f in all_fonts if is_collection_font(f)]
    
    print()
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"[RESULT] Total variants extracted: {total_extracted}")
    print(f"[OUTPUT] Location: {source_output_dir}")
    
    if len(remaining_collections) > 0:
        print(f"\n[WARN] {len(remaining_collections)} collection font(s) still remain:")
        for font in remaining_collections:
            rel_path = font.relative_to(source_output_dir)
            print(f"    - {rel_path}")
        print("\n[INFO] You may need to run the script again or check these fonts manually.")
    else:
        print(f"\n[SUCCESS] All fonts have been fully extracted to single fonts.")
        print("[OK] No collection fonts remain - all fonts are now individual files.")
    
    print("=" * 80)
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()

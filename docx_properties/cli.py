import argparse
import sys
from .core import DocxProperties, PropertyValue

def main():
    parser = argparse.ArgumentParser(description='Edit Word document properties')
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # get command (deprecated)
    get_parser = subparsers.add_parser('get', help='Read property (deprecated)')
    get_parser.add_argument('file', help='Word document (.docx)')
    get_parser.add_argument('property', nargs='?', help='Property name (optional)')
    get_parser.add_argument('--value', action='store_true', help='Show only value')
    get_parser.add_argument('--type', action='store_true', help='Show only type')
    
    # get-custom command
    get_custom_parser = subparsers.add_parser('get-custom', help='Show custom properties')
    get_custom_parser.add_argument('file', help='Word document (.docx)')
    get_custom_parser.add_argument('property', nargs='?', help='Custom property name (optional)')
    get_custom_parser.add_argument('--value', action='store_true', help='Show only value')
    get_custom_parser.add_argument('--type', action='store_true', help='Show only type')
    
    # get-core command
    get_core_parser = subparsers.add_parser('get-core', help='Show core properties')
    get_core_parser.add_argument('file', help='Word document (.docx)')
    get_core_parser.add_argument('property', nargs='?', help='Core property name (optional)')
    get_core_parser.add_argument('--value', action='store_true', help='Show only value')
    get_core_parser.add_argument('--type', action='store_true', help='Show only type')
    
    # get-all command
    get_all_parser = subparsers.add_parser('get-all', help='Show all properties')
    get_all_parser.add_argument('file', help='Word document (.docx)')
    get_all_parser.add_argument('--value', action='store_true', help='Show only values')
    get_all_parser.add_argument('--type', action='store_true', help='Show only types')
    
    # set command (deprecated)
    set_parser = subparsers.add_parser('set', help='Set property (deprecated)')
    set_parser.add_argument('file', help='Word document (.docx)')
    set_parser.add_argument('property', help='Property name')
    set_parser.add_argument('value', help='Property value')
    set_parser.add_argument('--no-update', action='store_true', help='Do not update fields')
    
    # set-custom command
    set_custom_parser = subparsers.add_parser('set-custom', help='Set custom property')
    set_custom_parser.add_argument('file', help='Word document (.docx)')
    set_custom_parser.add_argument('property', help='Custom property name')
    set_custom_parser.add_argument('value', help='Property value')
    set_custom_parser.add_argument('--no-update', action='store_true', help='Do not update fields')
    
    # update command
    update_parser = subparsers.add_parser('update', help='Update fields in document')
    update_parser.add_argument('file', help='Word document (.docx)')
    
    # filter-custom command
    filter_custom_parser = subparsers.add_parser('filter-custom', help='Filter custom properties')
    filter_custom_parser.add_argument('file', help='Word document (.docx)')
    filter_custom_parser.add_argument('--type', dest='filter_type', help='Filter by type (e.g. boolean, string)')
    filter_custom_parser.add_argument('--value', dest='filter_value', help='Filter by value')

    # filter-core command
    filter_core_parser = subparsers.add_parser('filter-core', help='Filter core properties')
    filter_core_parser.add_argument('file', help='Word document (.docx)')
    filter_core_parser.add_argument('--type', dest='filter_type', help='Filter by type (e.g. datetime, string)')
    filter_core_parser.add_argument('--value', dest='filter_value', help='Filter by value')

    # filter-all command
    filter_all_parser = subparsers.add_parser('filter-all', help='Filter all properties')
    filter_all_parser.add_argument('file', help='Word document (.docx)')
    filter_all_parser.add_argument('--type', dest='filter_type', help='Filter by type')
    filter_all_parser.add_argument('--value', dest='filter_value', help='Filter by value')
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        sys.exit(1)
    
    try:
        doc = DocxProperties(args.file)
        
        if args.command == 'get':
            if args.property:
                prop = doc.get_property(args.property)
                if prop is not None:
                    if args.type:
                        print(f"{args.property}: {prop.type}")
                    elif args.value:
                        print(f"{args.property}: {prop.value}")
                    else:
                        print(f"{args.property}: {prop.value} (Type: {prop.type})")
                else:
                    print(f"Property '{args.property}' not found.")
            else:
                # Show custom properties by default
                properties = doc.get_custom_properties()
                if properties:
                    print("Custom Properties:")
                    print("------------------")
                    for name, prop in properties.items():
                        if args.type:
                            print(f"{name}: {prop.type}")
                        elif args.value:
                            print(f"{name}: {prop.value}")
                        else:
                            print(f"{name}: {prop.value} (Type: {prop.type})")
                else:
                    print("No custom properties found.")
        
        elif args.command == 'get-custom':
            if args.property:
                prop = doc.get_custom_property(args.property)
                if prop is not None:
                    if args.type:
                        print(f"{args.property}: {prop.type}")
                    elif args.value:
                        print(f"{args.property}: {prop.value}")
                    else:
                        print(f"{args.property}: {prop.value} (Type: {prop.type})")
                else:
                    print(f"Custom property '{args.property}' not found.")
            else:
                properties = doc.get_custom_properties()
                if properties:
                    print("Custom Properties:")
                    print("------------------")
                    for name, prop in properties.items():
                        if args.type:
                            print(f"{name}: {prop.type}")
                        elif args.value:
                            print(f"{name}: {prop.value}")
                        else:
                            print(f"{name}: {prop.value} (Type: {prop.type})")
                else:
                    print("No custom properties found.")
                
        elif args.command == 'get-core':
            if args.property:
                prop = doc.get_core_property(args.property)
                if prop is not None:
                    if args.type:
                        print(f"{args.property}: {prop.type}")
                    elif args.value:
                        print(f"{args.property}: {prop.value}")
                    else:
                        print(f"{args.property}: {prop.value} (Type: {prop.type})")
                else:
                    print(f"Core property '{args.property}' not found.")
            else:
                properties = doc.get_core_properties()
                if properties:
                    print("Core Properties:")
                    print("---------------")
                    for name, prop in properties.items():
                        if args.type:
                            print(f"{name}: {prop.type}")
                        elif args.value:
                            print(f"{name}: {prop.value}")
                        else:
                            print(f"{name}: {prop.value} (Type: {prop.type})")
                else:
                    print("No core properties found.")
                
        elif args.command == 'get-all':
            properties = doc.get_all_properties()
            if properties:
                print("All Properties:")
                print("--------------")
                for name, prop in properties.items():
                    if args.type:
                        print(f"{name}: {prop.type}")
                    elif args.value:
                        print(f"{name}: {prop.value}")
                    else:
                        print(f"{name}: {prop.value} (Type: {prop.type})")
            else:
                print("No properties found.")
        
        elif args.command in ['set', 'set-custom']:
            # Interpret value
            value = args.value.lower()
            if value in ["true", "yes", "1"]:
                value = True
            elif value in ["false", "no", "0"]:
                value = False
            else:
                try:
                    if '.' in value:
                        value = float(value)
                    else:
                        value = int(value)
                except:
                    value = args.value
            
            if args.command == 'set':
                success = doc.set_property(args.property, value, not args.no_update)
            else:
                success = doc.set_custom_property(args.property, value, not args.no_update)
                
            if success:
                print(f"Property {args.property} was set to {value}.")
            else:
                print(f"Error setting property {args.property}.")
                sys.exit(1)
        
        elif args.command == 'update':
            success = doc.update_fields()
            if success:
                print(f"Fields in {args.file} were updated.")
            else:
                print(f"Fields could not be updated. Is pywin32 installed?")
                sys.exit(1)
                
        elif args.command.startswith('filter-'):
            # Interpret value for filter if specified
            filter_value = None
            if args.filter_value:
                value = args.filter_value.lower()
                if value in ["true", "yes", "1"]:
                    filter_value = True
                elif value in ["false", "no", "0"]:
                    filter_value = False
                else:
                    try:
                        if '.' in value:
                            filter_value = float(value)
                        else:
                            filter_value = int(value)
                    except:
                        filter_value = args.filter_value
            
            # Call the appropriate filter method
            if args.command == 'filter-custom':
                properties = doc.filter_custom_properties(args.filter_type, filter_value)
                title = "Filtered custom properties"
            elif args.command == 'filter-core':
                properties = doc.filter_core_properties(args.filter_type, filter_value)
                title = "Filtered core properties"
            else:  # filter-all
                properties = doc.filter_all_properties(args.filter_type, filter_value)
                title = "Filtered properties"
            
            # Display results
            print(f"{title}:")
            print("-" * len(title))
            
            if properties:
                for name, prop in properties.items():
                    print(f"{name}: {prop.value} (Type: {prop.type})")
            else:
                if args.filter_type and filter_value is not None:
                    print(f"No properties found with type '{args.filter_type}' and value '{filter_value}'.")
                elif args.filter_type:
                    print(f"No properties found with type '{args.filter_type}'.")
                elif filter_value is not None:
                    print(f"No properties found with value '{filter_value}'.")
                else:
                    print("No properties found.")
    
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
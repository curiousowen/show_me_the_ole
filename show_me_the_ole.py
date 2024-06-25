import sys
import struct

def print_usage():
    print("Usage: python show_me_the_ole.py <ole_file>")
    print("Example: python show_me_the_ole.py example.doc")

def validate_file(file_path):
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
            if len(header) != 8 or header[:4] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                print(f"Error: The file '{file_path}' is not a valid OLE file.")
                sys.exit(1)
    except IOError:
        print(f"Error: The file '{file_path}' does not exist.")
        sys.exit(1)

def read_sector(f, sector, sector_size):
    f.seek((sector + 1) * sector_size)
    return f.read(sector_size)

def list_streams(file_path):
    sector_size = 512
    with open(file_path, 'rb') as f:
        header = f.read(sector_size)
        
        if header[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            print("Not a valid OLE file.")
            return
        
        # Number of sectors in the SAT
        num_sectors_in_sat = struct.unpack('<I', header[0x2C:0x30])[0]
        
        # SAT (Sector Allocation Table) sectors
        sat_sectors = [struct.unpack('<I', header[0x4C + 4 * i: 0x50 + 4 * i])[0] for i in range(109)]
        
        # Directory Entry Sector
        dir_sector = struct.unpack('<I', header[0x30:0x34])[0]
        
        print("OLE Header and Sector Allocation Table (SAT):")
        print(f"Number of SAT sectors: {num_sectors_in_sat}")
        print(f"SAT sectors: {sat_sectors}")
        
        print("\nStreams and Storage objects:")
        # Reading the Directory Entry
        sector_size = 512
        sector = dir_sector
        while sector != 0xFFFFFFFE:
            data = read_sector(f, sector, sector_size)
            for i in range(0, len(data), 128):
                dir_entry = data[i:i + 128]
                if len(dir_entry) < 128:
                    continue
                name_length = struct.unpack('<H', dir_entry[0x40:0x42])[0]
                if name_length == 0:
                    continue
                name = dir_entry[:name_length-2].decode('utf-16le', errors='ignore')
                obj_type = dir_entry[0x42]
                
                if obj_type == 1:
                    obj_type_str = "Storage"
                elif obj_type == 2:
                    obj_type_str = "Stream"
                else:
                    obj_type_str = "Unknown"
                
                print(f"Name: {name}, Type: {obj_type_str}")
            sector = struct.unpack('<I', data[sector_size - 4:])[0]

def main():
    if len(sys.argv) != 2:
        print("Error: Incorrect number of arguments.")
        print_usage()
        sys.exit(1)

    file_path = sys.argv[1]
    validate_file(file_path)
    
    list_streams(file_path)

if __name__ == "__main__":
    main()

import barcode
from barcode.writer import ImageWriter
from barcode import Code128



# Function to generate barcode for each remote
def generate_barcode(remote_id):
    code128 = barcode.get_barcode_class('code128')  # Using Code128 format
    barcode_instance = code128(remote_id, writer=ImageWriter())
    barcode_instance.save(f"remote_{remote_id}")  # Saves the barcode as an image file

# Generate barcodes for 100 remotes
def generate_all_barcodes():
    for i in range(1, 40):
        remote_id = f"SP{i:04d}"  # Creates a remote ID like R0001, R0002, ...
        generate_barcode(remote_id)
        print(f"Barcode for {remote_id} generated.")

if __name__ == "__main__":
    generate_all_barcodes()

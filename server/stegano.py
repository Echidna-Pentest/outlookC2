from PIL import Image

def encode_message(image_path, output_path, message):
    """Encode a message into an image."""
    img = Image.open(image_path)
    encoded = img.copy()
    width, height = img.size
    pixels = encoded.load()

    # Convert the message to binary and append a termination marker
    binary_message = ''.join(format(ord(char), '08b') for char in message) + '1111111111111110'

    binary_index = 0
    for y in range(height):
        for x in range(width):
            if binary_index < len(binary_message):
                pixel = list(pixels[x, y])

                # Modify the LSB of each color channel
                for channel in range(3):  # Red, Green, Blue
                    if binary_index < len(binary_message):
                        pixel[channel] = pixel[channel] & ~1 | int(binary_message[binary_index])
                        binary_index += 1
                
                pixels[x, y] = tuple(pixel)

    encoded.save(output_path)
    print(f"Message encoded and saved to {output_path}")


def decode_message(image_path):
    """Decode a message from an image."""
    img = Image.open(image_path)
    pixels = img.load()
    width, height = img.size

    binary_message = ''
    for y in range(height):
        for x in range(width):
            pixel = pixels[x, y]

            # Extract the LSB of each color channel
            for channel in range(3):  # Red, Green, Blue
                binary_message += str(pixel[channel] & 1)

    # Convert binary to characters until the termination marker
    message = ''
    for i in range(0, len(binary_message), 8):
        byte = binary_message[i:i+8]
        if byte == '11111111':  # Adjusted termination marker (8 bits instead of 16)
            break
        if len(byte) < 8:  # Ignore incomplete bytes
            continue
        message += chr(int(byte, 2))

    return message


# Example usage
if __name__ == "__main__":
    # Encode a message
    original_image = "original.png"  # Replace with your image path
    output_image = "encoded_image.png"     # Path for saving the encoded image
    secret_message = "whoami; ls; ipconfig;"

    encode_message(original_image, output_image, secret_message)

    # Decode the message
    decoded_message = decode_message(output_image)
    print("Decoded message:", decoded_message)

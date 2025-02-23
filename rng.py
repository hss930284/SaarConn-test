import uuid

def generate_uuid():
    # Generate a random UUID
    random_uuid = uuid.uuid4()
    # Convert the UUID to a string in the specified format
    uuid_str = str(random_uuid)
    return uuid_str

# Example usage
if __name__ == "__main__":
    unique_id = generate_uuid()
    print(unique_id)
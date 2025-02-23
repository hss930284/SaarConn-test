import sys
import logging

class DataProcessor:
    def __init__(self):
        # Configure logging
        logging.basicConfig(level=logging.ERROR)

    def is_boolean(self, obj, depth=0):
        """
        Check if the object is of boolean type or contains a boolean in nested lists/tuples.
        
        Args:
            obj: The object to check.
            depth: Current recursion depth to avoid stack overflow.
        
        Returns:
            bool: True if the object is a boolean or contains a boolean, False otherwise.
        
        Raises:
            RecursionError: If maximum recursion depth is exceeded.
        """
        # Check recursion depth to avoid stack overflow
        if depth > sys.getrecursionlimit():
            raise RecursionError("Maximum recursion depth exceeded")
        
        # Validate input
        if obj is None:
            return False
        if not hasattr(obj, '__class__'):
            return False
        
        # Check if the object is of boolean type
        if "'bool'" in str(obj.__class__):
            return True
        # If the object is a non-empty list or tuple, check the first element recursively
        elif "'list'" in str(obj.__class__) and bool(obj):
            return self.is_boolean(obj[0], depth + 1)
        elif "'tuple'" in str(obj.__class__) and bool(obj):
            return self.is_boolean(obj[0], depth + 1)
        else:
            return False  # Return False if the object is not a boolean

    def value_to_str(self, value):
        """
        Convert a value to its string representation.
        
        Args:
            value: The value to convert.
        
        Returns:
            str or list: The string representation of the value or a list of string representations.
        
        Raises:
            TypeError: If the value is of an unsupported type.
        """
        if value is None:
            return ''
        elif isinstance(value, list):
            return [self.value_to_str(v) for v in value]
        elif self.is_boolean(value):
            # If the value is boolean, return "1" for True and "0" for False
            return "1" if value else "0"
        elif isinstance(value, int):
            return str(int(value))
        elif isinstance(value, float):
            return str(value)
        elif isinstance(value, str):
            return value
        else:
            # Raise an error for unsupported types
            raise TypeError(f"Unsupported type: {type(value)}")

    def read_file(self, file_path):
        """
        Read data from a file with error handling.
        
        Args:
            file_path (str): The path to the file to read.
        
        Returns:
            str: The content of the file.
        
        Raises:
            FileNotFoundError: If the file does not exist.
            IOError: If an error occurs while reading the file.
            Exception: For any other unexpected errors.
        """
        try:
            with open(file_path, 'r') as file:
                data = file.read()
            return data
        except FileNotFoundError:
            logging.error("File not found.")
            raise
        except IOError as e:
            logging.error(f"An error occurred while reading the file: {e}")
            raise
        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")
            raise
from colorama import Fore, Style, init

init(autoreset=True)

# Create a dictionary of common colors
color_dict = {
    'RED': Fore.RED,
    'GREEN': Fore.GREEN,
    'BLUE': Fore.BLUE,
    'YELLOW': Fore.YELLOW,
    'MAGENTA': Fore.MAGENTA,
    'CYAN': Fore.CYAN,
    'WHITE': Fore.WHITE,
    'BLACK': Fore.BLACK,
    'RESET': Fore.RESET  # To reset the color back to default
}

# Example usage
def color_examples():
    for color_name, color_code in color_dict.items():
        print(f"{color_code}{color_name} color example")

def color(string:str, color:str) -> str:
    colored_str = f"{color_dict[f'{color}']} {string}"
    return colored_str

def print_blue(string:str) -> None:
    print(f"{color_dict['BLUE']} {string}")

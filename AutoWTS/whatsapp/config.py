import os


user_name = os.getlogin()
CHROME_PROFILE_PATH = f'user-data-dir=C:\\Users\\{user_name}\\AppData\\Local\\Google\\Chrome\\User Data\\Whatsapp'
print(CHROME_PROFILE_PATH)

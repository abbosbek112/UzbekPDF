import os
import time
import shutil

# Tozalanishi kerak bo'lgan papka va vaqt (24 soat)
TEMP_DIR = os.path.join(os.path.dirname(__file__), 'temp')
AGE_LIMIT_SECONDS = 24 * 60 * 60

def cleanup_temp_dir():
    if not os.path.exists(TEMP_DIR):
        print(f"Papkasi mavjud emas: {TEMP_DIR}")
        return

    current_time = time.time()
    deleted_count = 0

    for root, dirs, files in os.walk(TEMP_DIR, topdown=False):
        for name in files:
            file_path = os.path.join(root, name)
            # Check file age
            if current_time - os.path.getmtime(file_path) > AGE_LIMIT_SECONDS:
                try:
                    os.remove(file_path)
                    print(f"O'chirildi (fayl): {file_path}")
                    deleted_count += 1
                except Exception as e:
                    print(f"Xato (fayl): {file_path} - {e}")
        
        for name in dirs:
            dir_path = os.path.join(root, name)
            # Remove empty directories if they are older than AGE_LIMIT_SECONDS
            if current_time - os.path.getmtime(dir_path) > AGE_LIMIT_SECONDS:
                try:
                    # check if empty
                    if not os.listdir(dir_path):
                        os.rmdir(dir_path)
                        print(f"O'chirildi (bo'sh papka): {dir_path}")
                        deleted_count += 1
                    else:
                        shutil.rmtree(dir_path)
                        print(f"O'chirildi (to'la papka): {dir_path}")
                        deleted_count += 1
                except Exception as e:
                    print(f"Xato (papka): {dir_path} - {e}")
                    
    print(f"Tozalash tugadi. Jami {deleted_count} ta element o'chirildi.")

if __name__ == "__main__":
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Tozalash jarayoni boshlandi...")
    cleanup_temp_dir()

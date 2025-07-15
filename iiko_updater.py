import configparser
import subprocess
import os
import shutil
import ftplib
import threading
import zipfile
import time
import re
import sys
from urllib.parse import urlparse
from tqdm import tqdm
import ctypes # Для проверки прав администратора и UAC
from ctypes import wintypes # Для проверки прав администратора

# --- Константы и настройки ---
CONFIG_FILE = 'config.ini'
DEFAULT_CONFIG = {
    'FTP': {'url': ''}, 
    'SMB': {'path': ''}, 
    'Services': {'keywords': 'Tomcat,iiko,RMS,ChainServer,iikoChain'}, 
    'General': {'download_dir': './downloads', 'backup_dir': './backups'}
}
REQUIRED_BACKUP_FOLDERS = ['exploded', 'tools', 'tomcat9', 'logs']
FOLDERS_TO_REPLACE = ['exploded', 'tools', 'tomcat9']
LOG_FILE_NAME = 'startup.log'
SUCCESS_LOG_MESSAGE = 'STARTED_SUCCESSFULLY'
SERVICE_POLL_INTERVAL = 2 # seconds
SERVICE_TIMEOUT = 600 # seconds
LOG_POLL_INTERVAL = 1 # seconds
LOG_TIMEOUT = 900 # seconds (15 minutes)

try:
    from colorama import Fore, Style, init

    # Определяем наши цвета для удобства
    COLOR_MAP = {
    'header': Fore.WHITE,                  # Чистый белый для заголовков
    'prompt': Fore.YELLOW,                 # Классический желтый (без яркости) для ввода
    'success': Fore.GREEN,                 # Стандартный зеленый для успеха
    'error': Fore.RED,                     # Стандартный красный для ошибок
    'warning': Fore.YELLOW,                # Желтый для предупреждений (как и для ввода)
    'info': Fore.CYAN,                     # Спокойный голубой (Cyan) для информации
    'step': Fore.CYAN,                     # Этот же голубой для обозначения шагов
    }

    def cprint(text, color=''):
        """Выводит цветной текст."""
        color_code = COLOR_MAP.get(color, '')
        print(f"{color_code}{text}{Style.RESET_ALL}")

    def prompt(text):
        """Запрашивает ввод у пользователя с цветным текстом."""
        color_code = COLOR_MAP.get('prompt', '')
        # Для prompt мы не можем использовать autoreset, поэтому сбрасываем стиль вручную
        return input(f"{color_code}{text}{Style.RESET_ALL}")
    
    def start_color(color=''):
        """Включает цвет для последующих print, не добавляя перевод строки."""
        color_code = COLOR_MAP.get(color, '')
        # Мы используем print, чтобы colorama перехватила вызов, но end='' предотвращает перевод строки.
        print(color_code, end='')

    def end_color():
        """Сбрасывает цвет к стандартному, не добавляя перевод строки."""
        print(Style.RESET_ALL, end='')

except ImportError:
    # Если colorama не установлена, создаем "пустые" функции, чтобы скрипт не сломался.
    print("ПРЕДУПРЕЖДЕНИЕ: Библиотека 'colorama' не найдена. Вывод будет монохромным.")
    print("Для цветного вывода выполните: pip install colorama")
    def cprint(text, color=''):
        print(text)
    def prompt(text):
        return prompt(text)
    def start_color(color=''): pass
    def end_color(): pass

# --- Вспомогательные функции ---

def is_admin():
    """Проверяет, запущен ли скрипт от имени администратора (только для Windows)."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def elevate_privileges():
    """Пытается перезапустить текущий скрипт с повышенными правами UAC."""
    if not is_admin():
        cprint("\nПопытка повышения прав (требуется подтверждение UAC)...", 'step')
        python_exe = sys.executable # Путь к текущему исполняемому файлу Python
        script_path = os.path.abspath(sys.argv[0]) # Путь к текущему скрипту
        
        # Собираем аргументы, с которыми был запущен скрипт
        script_args = ' '.join([f'"{arg}"' if ' ' in arg else arg for arg in sys.argv[1:]]) # Экранируем аргументы с пробелами

        # Команда для ShellExecuteW: (hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd)
        # lpOperation="runas" для повышения прав
        # lpFile - что запускаем (python.exe)
        # lpParameters - аргументы для lpFile (script_path и его аргументы)
        # lpDirectory - рабочая директория
        # nShowCmd = 1 (SW_SHOWNORMAL) для обычного окна
        
        # Важно: lpParameters должен быть одной строкой.
        # Если путь к скрипту содержит пробелы, его тоже нужно взять в кавычки.
        full_command_line = f'"{script_path}" {script_args}' if ' ' in script_path else f'{script_path} {script_args}'

        try:
            # ShellExecuteW возвращает HANDLE процесса (если > 32) или код ошибки (0-32)
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", python_exe, full_command_line, None, 1
            )
            if ret <= 32:
                cprint(f"!!! Ошибка или отмена повышения прав. Код ошибки ShellExecuteW: {ret}", 'error')
                prompt("Нажмите Enter для выхода.")
                sys.exit(1) # Завершаем текущий процесс с ошибкой
            else:
                cprint("Попытка перезапуска с повышенными правами. Текущий процесс будет закрыт.", 'warning')
                sys.exit(0) # Завершаем текущий процесс успешно, повышенный процесс возьмет управление
        except Exception as e:
            cprint(f"!!! Неизвестная ошибка при попытке повышения прав: {e}", 'error')
            prompt("Нажмите Enter для выхода.")
            sys.exit(1)


def load_config(filename):
    """Загружает конфигурацию из файла."""
    config = configparser.ConfigParser()
    if not os.path.exists(filename):
        cprint(f"Файл конфигурации '{filename}' не найден. Создаю файл с настройками по умолчанию.", 'warning')
        for section, options in DEFAULT_CONFIG.items():
            config.add_section(section)
            for key, value in options.items():
                config.set(section, key, value)
        with open(filename, 'w') as configfile:
            config.write(configfile)
        cprint(f"Создан файл '{filename}'. Пожалуйста, отредактируйте его.", 'warning')
        return None 
    config.read(filename)
    return config

def decode_bytes_with_fallbacks(byte_data, fallbacks=['cp866', 'cp1251', 'utf-8', sys.stdout.encoding], errors='ignore'):
    """Пытается декодировать байты, перебирая кодировки, возвращает первый успешный результат."""
    if not isinstance(byte_data, bytes): return byte_data
    if not byte_data: return "" 
    for encoding in fallbacks:
        if not encoding: continue
        try: return byte_data.decode(encoding, errors=errors)
        except Exception: pass 
    return None 


def run_command(command, shell=True, check=True, text=False, encoding=None, errors='ignore'):
    """Выполняет команду, получает вывод (байты или текст) и пытается его декодировать при text=False."""
    command_str = ' '.join(command) if isinstance(command, list) else command
    try:
        if text:
            result = subprocess.run(command, capture_output=True, text=True, shell=shell, check=check, encoding=encoding, errors=errors)
            stdout_output = result.stdout
            stderr_output = result.stderr 
        else:
            result = subprocess.run(command, capture_output=True, shell=shell, check=check)
            stdout_output = decode_bytes_with_fallbacks(result.stdout)
            stderr_output = decode_bytes_with_fallbacks(result.stderr, fallbacks=[sys.stderr.encoding, 'cp866', 'utf-8'], errors='ignore')

        if result.returncode != 0 and stderr_output:
            cprint(f"!!! Ошибка выполнения команды '{command_str}':", 'error')
            cprint(f"!!! Stderr:\n{stderr_output}\n!!!", 'error')
            return None 
        return stdout_output
            
    except subprocess.CalledProcessError as e:
        cprint(f"\n!!! Ошибка выполнения команды: {e}", 'error')
        stderr_output = decode_bytes_with_fallbacks(e.stderr, fallbacks=[sys.stderr.encoding, 'cp866', 'utf-8'], errors='ignore')
        cprint(f"!!! Stderr:\n{stderr_output if stderr_output else '!!! ОШИБКА ДЕКОДИРОВАНИЯ STDOUT ИЛИ ПУСТОЙ STDOUT'}\n!!!", 'error')
        return None
    except FileNotFoundError:
        cmd_name = command[0] if isinstance(command, list) else command.split()[0]
        cprint(f"\n!!! Ошибка: Команда '{cmd_name}' не найдена. Убедитесь, что она доступна в PATH.", 'error')
        return None
    except Exception as e:
        cmd_name = command[0] if isinstance(command, list) else command.split()[0]
        cprint(f"\n!!! Неизвестная ошибка при выполнении команды '{cmd_name}': {e}", 'error')
        return None

def get_service_info_wmic(keywords):
    """Находит службы iikoRMS/iikoChain по ключевым словам, используя wmic."""
    cprint("\n--- Поиск служб iikoRMS/iikoChain через wmic ---", 'step')
    services_info = []
    
    output = run_command('wmic service get Name,DisplayName,State,PathName /format:list', shell=True, check=False, text=False)
    if not output:
        cprint("Не удалось получить или декодировать список служб.", 'error')
        return services_info

    raw_blocks = re.split(r'\r\r\n\r\r\n', output.strip())
    
    keyword_list = [k.strip().lower() for k in keywords.split(',') if k.strip()]

    for raw_block in raw_blocks:
        block = raw_block.strip()
        if not block: continue

        current_service_data = {}
        lines = block.split('\r\r\n')
        
        for line in lines:
            line = line.strip()
            if not line or '=' not in line: continue
            try:
                key, value = line.split('=', 1)
                current_service_data[key.strip()] = value.strip()
            except Exception: continue

        name = current_service_data.get('Name')
        display = current_service_data.get('DisplayName')
        state = current_service_data.get('State')
        path_raw = current_service_data.get('PathName')
        
        if not name or not display or state is None or path_raw is None: continue
        if not any(k in name.lower() or k in display.lower() for k in keyword_list): continue

        binary_path = None
        if path_raw.startswith('"'):
            m = re.match(r'^"(.+?)"', path_raw)
            binary_path = m.group(1) if m else None
        else:
            binary_path = path_raw.split(' ')[0]

        if not binary_path: continue
        binary_path = os.path.abspath(binary_path)
        if not os.path.isfile(binary_path): continue

        try:
            exe_dir = os.path.dirname(binary_path)
            tomcat_dir = os.path.dirname(exe_dir)
            server_dir = os.path.dirname(tomcat_dir)

            # Проверяем, что рассчитанная директория реально существует и является директорией
            # А также, что внутри нее есть ожидаемая структура iiko (Tomcat9/bin/tomcat9.exe)
            expected_iiko_exe_path = os.path.join(server_dir, 'Tomcat9', 'bin', os.path.basename(binary_path))
            
            if os.path.isdir(server_dir) and os.path.exists(expected_iiko_exe_path) and os.path.isfile(expected_iiko_exe_path):
                product_type = 'iikoChain' if 'chain' in name.lower() or 'chain' in display.lower() else 'iikoRMS'
                cprint(f"Найден сервер: '{display}' ({name}, {state}) в директории '{server_dir}'", 'success')
                services_info.append({
                    'name': name,
                    'display_name': display,
                    'status': state,
                    'binary_path': binary_path,
                    'server_dir': server_dir,
                    'type': product_type 
                })
        except Exception as e:
             cprint(f"!!! Ошибка при определении директории сервера из пути '{binary_path}': {e}. Пропускаю.", 'warning')

    if not services_info:
        cprint("Серверы iikoRMS/iikoChain не найдены по указанным ключевым словам и путям.", 'warning')

    return services_info

def get_service_status_wmic(service_name):
    """Получает текущий статус службы, используя wmic."""
    output = run_command(f'wmic service where Name="{service_name}" get State /value', shell=True, check=False, text=False)
    if output:
        state_match = re.search(r'^State=(.+)$', output.strip(), re.MULTILINE)
        if state_match: return state_match.group(1).strip().capitalize() 
    return None

def get_service_logon_account_wmic(service_name):
    """Получает учетную запись службы, используя wmic."""
    output = run_command(f'wmic service where Name="{service_name}" get StartName /value', shell=True, check=False, text=False)
    if output:
        account_match = re.search(r'^StartName=(.+)$', output.strip(), re.MULTILINE)
        if account_match: return account_match.group(1).strip()
    return None

def set_service_logon_account_sc(service_name, account='LocalSystem', username=None, password=None):
    """Изменяет учетную запись службы на 'LocalSystem' с использованием sc.exe config."""
    cprint(f"Попытка изменить учетную запись службы '{service_name}' на '{account}' (через sc.exe config)...", 'step')
    
    if account.lower() == 'localsystem':
        command = f'sc.exe config "{service_name}" obj= "LocalSystem" password= ""'
    else:
        if not username or not password:
            cprint("!!! Ошибка: Для учетной записи, отличной от LocalSystem, необходимо указать имя пользователя и пароль.", 'error')
            return False
        command = f'sc.exe config "{service_name}" obj= "{username}" password= "{password}"'

    result = run_command(command, shell=True, check=False, text=False) 
    
    if result is None: 
        cprint(f"!!! Ошибка при изменении учетной записи службы '{service_name}'.", 'error')
        return False
    
    if "FAILED" in result.upper():
        cprint(f"!!! sc.exe config вернул FAILED при изменении учетной записи службы: {result}", 'error')
        return False

    cprint("Учетная запись службы успешно изменена.", 'success')
    return True


def wait_for_service_status(service_name, target_status, timeout=SERVICE_TIMEOUT, poll_interval=SERVICE_POLL_INTERVAL):
    """Ожидает перехода службы в заданный статус, используя wmic."""
    start_color('info')
    print(f"Ожидание статуса '{target_status}' для службы '{service_name}'...", end='', flush=True)
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_status = get_service_status_wmic(service_name) 
        if current_status is None:
             cprint(f"\nОшибка: Служба '{service_name}' не найдена или недоступна.", 'error')
             return False
        if current_status == target_status:
            cprint(" Готово.", 'success')
            return True
        print(".", end='', flush=True)
        time.sleep(poll_interval)
    end_color()
    cprint("\nТаймаут ожидания статуса службы.", 'warning')
    return False

def start_service_ps(service_name):
     """Запускает службу с использованием PowerShell."""
     cprint(f"Попытка запуска службы '{service_name}' (через PowerShell)...", 'step')
     return run_command(['powershell', '-Command', f"Start-Service -Name '{service_name}'"], shell=False, check=False, text=True, encoding='utf-8') is not None 

def stop_service_ps(service_name):
     """Останавливает службу с использованием PowerShell."""
     cprint(f"Попытка остановки службы '{service_name}' (через PowerShell)...", 'step')
     return run_command(['powershell', '-Command', f"Stop-Service -Name '{service_name}'"], shell=False, check=False, text=True, encoding='utf-8') is not None


def connect_ftp(ftp_url):
    """Подключается к FTP серверу."""
    if not ftp_url: return None, None
    try:
        normalized_url = ftp_url.replace('\\', '/')
        if normalized_url.startswith('ftp:/') and not normalized_url.startswith('ftp://'):
            # Заменяем только первое вхождение 'ftp:/' на 'ftp://'
            normalized_url = normalized_url.replace('ftp:/', 'ftp://', 1)
        parsed_url = urlparse(normalized_url)
        host = parsed_url.hostname
        port = parsed_url.port if parsed_url.port else 21
        username = parsed_url.username
        password = parsed_url.password
        cprint(f"Подключение к FTP: {host}:{port}...", 'step')
        ftp = ftplib.FTP()
        ftp.connect(host, port)
        ftp.login(username, password)
        cprint("Подключение успешно.", 'success')
        return ftp, parsed_url.path.lstrip('/')
    except ftplib.all_errors as e:
        cprint(f"Ошибка подключения к FTP: {e}", 'error')
        return None, None
    except Exception as e:
        cprint(f"Неизвестная ошибка при подключении к FTP: {e}", 'error')
        return None, None

def find_versions_on_ftp(ftp, base_path):
    """
    Ищет доступные версии обновлений на FTP, используя надежный метод проверки директорий.
    """
    cprint(f"Поиск доступных версий на FTP в директории /{base_path}...", 'step')
    available_versions = {}
    
    try:
        # 1. Переходим в базовую директорию для обновлений
        if base_path:
            try:
                ftp.cwd(base_path)
            except ftplib.all_errors as e:
                cprint(f"Ошибка при переходе в базовую директорию '{base_path}' на FTP: {e}", 'error')
                return {}

        # 2. Получаем список всех элементов (файлов и папок) в этой директории
        try:
            items_in_base_dir = ftp.nlst()
        except ftplib.all_errors as e:
            cprint(f"Не удалось получить список элементов в директории '{base_path}': {e}", 'error')
            return {}

        # 3. Итерируемся по каждому элементу, чтобы найти папки с версиями
        for version_dir in items_in_base_dir:
            # Пропускаем элементы, которые не похожи на номер версии
            if not re.match(r'^[\d\.]+$', os.path.basename(version_dir)):
                continue

            try:
                # 4. Проверяем, является ли элемент директорией, пытаясь в нее войти
                ftp.cwd(version_dir)
                
                # --- Если мы здесь, значит это директория ---
                
                # Получаем список архивов внутри папки версии
                files_in_dir = ftp.nlst()
                found_archives = [f for f in files_in_dir if f.lower().endswith('.zip') and ('chain' in f.lower() or 'rms' in f.lower())]
                
                if found_archives:
                    cprint(f"  Найдена версия '{os.path.basename(version_dir)}' на FTP с архивами: {', '.join(found_archives)}", 'success')
                    # Собираем полный путь для последующих операций
                    full_ftp_path = f"{base_path}/{os.path.basename(version_dir)}".strip('/')
                    unique_key = f"FTP {os.path.basename(version_dir)}"
                    available_versions[unique_key] = {
                        'source': 'ftp',
                        'version': os.path.basename(version_dir),
                        'path': full_ftp_path,
                        'archives': found_archives
                    }
                
                # 5. Обязательно возвращаемся на уровень выше
                ftp.cwd('..')

            except ftplib.error_perm:
                # Это ожидаемая ошибка, если `version_dir` - это файл, а не папка.
                # Просто игнорируем этот элемент и переходим к следующему.
                continue
            except Exception as e:
                cprint(f"  Неизвестная ошибка при обработке элемента '{version_dir}': {e}", 'warning')
                # Пытаемся вернуться на уровень выше на случай, если мы застряли где-то
                try: ftp.cwd('..')
                except: pass
                continue

    except Exception as e:
        cprint(f"Критическая ошибка при поиске версий на FTP: {e}", 'error')

    return available_versions

def find_versions_on_smb(smb_path):
    """Ищет доступные версии обновлений на SMB ресурсе."""
    if not smb_path: return {}
    cprint(f"Поиск доступных версий на SMB ресурсе в директории '{smb_path}'...", 'spep')
    available_versions = {}
    try:
        if not os.path.isdir(smb_path):
            cprint(f"Ошибка: SMB путь '{smb_path}' не является доступной директорией или недоступен.", 'error')
            return {}
        items = os.listdir(smb_path)
        directories = []
        for item in items:
             full_item_path = os.path.join(smb_path, item)
             if os.path.isdir(full_item_path): directories.append(item)

        for version_dir in directories:
            if not re.match(r'^[\d\.]+$', version_dir): continue 
            version_full_path = os.path.join(smb_path, version_dir)
            found_archives = []
            try:
                files_in_dir = os.listdir(version_full_path)
                found_archives = [f for f in files_in_dir if f.lower().endswith('.zip') and ('chain' in f.lower() or 'rms' in f.lower())]
            except OSError as e:
                cprint(f"  Ошибка доступа к директории '{version_full_path}' на SMB: {e}. Пропускаю.", 'warning')
                continue
            except Exception as e:
                cprint(f"  Неизвестная ошибка при доступе к директории '{version_full_path}' на SMB: {e}. Пропускаю.", 'warning')
                continue

            if found_archives:
                cprint(f"  Найдена версия '{version_dir}' на SMB с архивами: {', '.join(found_archives)}", 'success')
                unique_key = f"SMB {version_dir}"
                available_versions[unique_key] = {
                    'source': 'smb', 'version': version_dir, 'path': version_full_path, 'archives': found_archives 
                }
    except Exception as e:
        cprint(f"Неизвестная ошибка при поиске версий на SMB ресурсе: {e}", 'error')
    return available_versions


def download_update_archive(version_info, download_dir, ftp_conn=None):
    """Скачивает или копирует архив обновления в зависимости от источника."""
    source = version_info['source']
    source_path = version_info['path']
    archive_name = version_info['selected_archive'] 
    local_file_path = os.path.join(download_dir, archive_name)
    
    os.makedirs(download_dir, exist_ok=True)

    if source == 'ftp':
        if ftp_conn is None:
             cprint("Ошибка: FTP соединение не установлено для скачивания с FTP.", 'error')
             return None
        cprint(f"Скачивание '{archive_name}' из '{source_path}' (FTP) в '{local_file_path}'...", 'step')
        try:
            absolute_ftp_path = '/' + source_path.strip('/')
            ftp_conn.cwd(absolute_ftp_path)

            file_size = ftp_conn.size(archive_name)
            with open(local_file_path, 'wb') as local_file:
                with tqdm(
                    total=file_size,
                    unit='B', unit_scale=True, unit_divisor=1024,
                    desc=archive_name,
                    ascii=True # Для лучшей совместимости с консолями
                ) as progress:
                    def progress_callback(chunk):
                        local_file.write(chunk)
                        progress.update(len(chunk))
                    
                    ftp_conn.retrbinary(f'RETR {archive_name}', progress_callback)
            cprint("Скачивание завершено.", 'success')
            try:
                ftp_conn.cwd('/')
            except ftplib.all_errors:
                pass
            return local_file_path
        
        except Exception as e:
            cprint(f"Неизвестная ошибка при скачивании файла с FTP: {e}", 'error')
            try: ftp_conn.cwd('..')
            except: pass
            if os.path.exists(local_file_path): os.remove(local_file_path)
            return None

    elif source == 'smb':
        smb_file_path = os.path.join(source_path, archive_name)
        cprint(f"Копирование '{smb_file_path}' (SMB) в '{local_file_path}'...", 'step')
        try:
            file_size = os.path.getsize(smb_file_path)
            chunk_size = 1024 * 1024

            with open(smb_file_path, 'rb') as fsrc:
                with open(local_file_path, 'wb') as fdst:
                    with tqdm(
                        total=file_size,
                        unit='B', unit_scale=True, unit_divisor=1024,
                        desc=archive_name,
                        ascii=True
                    ) as progress:
                        while True:
                            chunk = fsrc.read(chunk_size)
                            if not chunk:
                                break
                            fdst.write(chunk)
                            progress.update(len(chunk))
            
            cprint("Копирование завершено.", 'success')
            return local_file_path
        except FileNotFoundError:
             cprint(f"Ошибка: Файл '{smb_file_path}' не найден.", 'error')
             return None
        except PermissionError:
             cprint(f"Ошибка: Нет прав для доступа к файлу '{smb_file_path}' или записи в '{local_file_path}'.", 'warning')
             return None
        except Exception as e:
            cprint(f"Неизвестная ошибка при копировании файла с SMB: {e}", 'error')
            if os.path.exists(local_file_path):
                 try: os.remove(local_file_path)
                 except: pass 
            return None
    else:
        cprint(f"Неизвестный источник обновления: {source}", 'error')
        return None


def backup_server_folders(server_dir, backup_base_dir):
    """
    Создает бэкап. Для папок 'exploded', 'tools', 'tomcat9' - перемещает их целиком.
    Для папки 'logs' - оставляет ее на месте, но перемещает ее содержимое в бэкап.
    """
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    backup_target_dir = os.path.join(backup_base_dir, f"iiko_backup_{timestamp}")
    
    cprint(f"Создание бэкапа в '{backup_target_dir}'...", 'step')
    
    # Предварительная очистка .exe файлов из папки 'exploded'
    # (Эта логика остается без изменений)
    print("  Предварительная очистка файлов установщиков из папки 'exploded'...")
    installers_to_clean = [
        os.path.join(server_dir, 'exploded', 'update', 'Front', 'Setup.Front.exe'),
        os.path.join(server_dir, 'exploded', 'update', 'BackOffice', 'Setup.RMS.BackOffice.exe'),
        os.path.join(server_dir, 'exploded', 'update', 'BackOffice', 'Setup.Chain.BackOffice.exe')
    ]
    for installer_path in installers_to_clean:
        if os.path.exists(installer_path):
            try:
                os.remove(installer_path)
                print(f"    Удален старый установщик: {os.path.basename(installer_path)}")
            except Exception as e:
                cprint(f"    ПРЕДУПРЕЖДЕНИЕ: Не удалось удалить файл '{installer_path}': {e}", 'warning')

    os.makedirs(backup_target_dir, exist_ok=True)
    
    # Список папок, которые были перемещены целиком (для отката)
    moved_folders = []
    try:
        for folder_name in REQUIRED_BACKUP_FOLDERS:
            src_path = os.path.join(server_dir, folder_name)
            
            if not os.path.exists(src_path):
                cprint(f"  Предупреждение: Папка '{folder_name}' не найдена. Пропускаю.", 'warning')
                continue

            if folder_name == 'logs':
                cprint(f"  Очистка папки '{folder_name}' (перемещение содержимого в бэкап)...", 'info')
                
                # Создаем папку 'logs' в директории бэкапа
                dest_logs_path = os.path.join(backup_target_dir, 'logs')
                os.makedirs(dest_logs_path, exist_ok=True)
                
                # Перемещаем каждый элемент из исходной папки logs в бэкап
                items_to_move = os.listdir(src_path)
                if not items_to_move:
                    cprint(f"    Папка '{folder_name}' пуста. Пропускаю.", 'info')
                    continue

                for item in items_to_move:
                    item_src_path = os.path.join(src_path, item)
                    item_dest_path = os.path.join(dest_logs_path, item)
                    shutil.move(item_src_path, item_dest_path)
                
                cprint(f"    Содержимое папки '{folder_name}' успешно перемещено. Сама папка сохранена.", 'success')
                # Важно: мы не добавляем 'logs' в список moved_folders, так как сама папка не перемещалась

            else:
                dest_path = os.path.join(backup_target_dir, folder_name)
                cprint(f"  Перемещение папки '{folder_name}'...", 'info')
                shutil.move(src_path, dest_path)
                moved_folders.append(folder_name) # Добавляем в список для отката
                cprint("    Успешно.", 'success')
        
        return backup_target_dir
        
    except Exception as e:
        cprint(f"!!! КРИТИЧЕСКАЯ ОШИБКА при создании бэкапа: {e}", 'error')
        cprint("!!! Попытка откатить перемещение...", 'warning')
        
        # Логика отката затронет только те папки, которые были перемещены целиком
        for folder_name in moved_folders:
            backup_folder_path = os.path.join(backup_target_dir, folder_name)
            original_path = os.path.join(server_dir, folder_name)
            try:
                if os.path.exists(backup_folder_path):
                    shutil.move(backup_folder_path, original_path)
                    cprint(f"  Папка '{folder_name}' возвращена на место.", 'success')
            except Exception as rollback_e:
                cprint(f"!!! НЕ УДАЛОСЬ вернуть папку '{folder_name}': {rollback_e}", 'error')
        return None

def perform_update(server_info, downloaded_archive_path, backup_base_dir, selected_version_info, ftp_conn=None):
    """
    Выполняет процедуру обновления с мониторингом лога в отдельном потоке
    и обработкой Ctrl+C.
    """
    # ... (весь код до Шага 5 остается без изменений) ...
    service_name = server_info['name']
    server_dir = server_info['server_dir']
    
    cprint(f"\n--- Запуск процедуры обновления для сервера '{server_info['display_name']}' ---", 'step')

    cprint("Шаг 1: Остановка службы...", 'step')
    if not stop_service_ps(service_name) or not wait_for_service_status(service_name, 'Stopped'):
        return False, None, False
    cprint("Служба успешно остановлена.", 'success')

    cprint("\nШаг 2: Создание бэкапа (перемещением)...", 'step')
    backup_path = backup_server_folders(server_dir, backup_base_dir)
    if not backup_path:
        start_service_ps(service_name)
        return False, None, False
    cprint(f"Бэкап создан в: {backup_path}", 'success')

    extract_dir = os.path.join(config['General']['download_dir'], 'extracted_update')
    
    try:
        cprint("\nШаг 3: Извлечение архива обновления...", 'step')
        if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(downloaded_archive_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        cprint("Архив успешно извлечен.", 'success')

        cprint("\nШаг 4: Развертывание новых файлов (перемещением)...", 'step')
        # ... (код развертывания) ...
        extracted_items = os.listdir(extract_dir)
        source_base_dir = extract_dir
        if len(extracted_items) == 1 and os.path.isdir(os.path.join(extract_dir, extracted_items[0])):
            source_base_dir = os.path.join(extract_dir, extracted_items[0])
            cprint(f"  Обнаружена корневая папка в архиве: '{extracted_items[0]}'. Работаем из нее.", 'info')
            
        # переносим кастомные .jar если есть
        migrate_custom_libs(
            backup_path=backup_path,
            source_base_dir=source_base_dir
        )
        for folder_name in FOLDERS_TO_REPLACE:
            src_path = os.path.join(source_base_dir, folder_name)
            dest_path = os.path.join(server_dir, folder_name)
            if os.path.exists(src_path):
                shutil.move(src_path, dest_path)
            else:
                raise FileNotFoundError(f"Папка '{folder_name}' не найдена в извлеченном архиве по пути '{src_path}'")
        cprint("Новые папки успешно развернуты.", 'success')

        cprint("\nШаг 4.5: Загрузка дополнительных установщиков...", 'step')
        if not download_and_place_installers(server_info, selected_version_info, ftp_conn):
            cprint("!!! ПРЕДУПРЕЖДЕНИЕ: Не удалось загрузить все необходимые установщики.", 'warning')
        else:
            cprint("Загрузка установщиков завершена.", 'success')

        cprint(f"\nШаг 5: Запуск службы '{service_name}'...", 'step')
        if not start_service_ps(service_name) or not wait_for_service_status(service_name, 'Running', timeout=90):
             raise RuntimeError("Служба не перешла в состояние 'Running' после обновления.")
        cprint("Служба запущена.", 'success')
        
        os.makedirs(os.path.join(server_dir, 'logs'), exist_ok=True)
        new_log_path = os.path.join(server_dir, 'logs', LOG_FILE_NAME)
        start_color('info')
        print(f"  Ожидание появления файла лога '{new_log_path}'...", end='', flush=True)
        start_time_wait_log = time.time()
        spinner = ['|', '/', '-', '\\']
        spinner_idx = 0
        while not os.path.exists(new_log_path):
            if time.time() - start_time_wait_log > SERVICE_TIMEOUT:
                print("\r" + " " * 60 + "\r", end="")
                raise TimeoutError(f"Таймаут ожидания файла лога '{new_log_path}'.")
            
            print(f"\r  Ожидание появления файла лога '{new_log_path}'... {spinner[spinner_idx]}", end='', flush=True)
            spinner_idx = (spinner_idx + 1) % len(spinner)
            time.sleep(LOG_POLL_INTERVAL)
        end_color()
        print("\r" + " " * 60 + "\r", end="")
        cprint("  Файл лога найден.", 'success')

        user_prompt = prompt("Выводить содержимое лога для контроля запуска? (да/нет/y/n): ").lower()
        if user_prompt not in ['да', 'y']:
            cprint("Мониторинг лога пропущен. Программа продолжит выполнение.", 'warning')
            return True, backup_path, True

        cprint("Мониторинг вывода лога... (Нажмите Ctrl+C для прерывания и перехода к загрузке бэкапа)", 'warning')
        stop_event = threading.Event()
        success_event = threading.Event()
        
        log_thread = threading.Thread(
            target=monitor_log_in_thread,
            args=(new_log_path, stop_event, success_event, SUCCESS_LOG_MESSAGE)
        )
        log_thread.start()
        
        start_time_monitor = time.time()
        monitoring_interrupted = False

        try:
            while log_thread.is_alive():
                # Проверяем, не сообщил ли поток об успехе
                if success_event.is_set():
                    cprint("\nОБНОВЛЕНИЕ УСПЕШНО ЗАВЕРШЕНО: Найдено сообщение об успехе.", 'success')
                    return True, backup_path, False # Успех, мониторинг НЕ пропущен

                # Проверяем таймаут
                if time.time() - start_time_monitor > LOG_TIMEOUT:
                    cprint("\n!!! Таймаут ожидания сообщения в логе.", 'warning')
                    raise TimeoutError("Таймаут ожидания сообщения в логе.")
                
                # Ждем 1 секунду. Это позволяет циклу быть отзывчивым к Ctrl+C
                time.sleep(1)

            if success_event.is_set():
                cprint("\nОБНОВЛЕНИЕ УСПЕШНО ЗАВЕРШЕНО: Найдено сообщение об успехе.", 'success')
                return True, backup_path, False
            
            raise RuntimeError("Неожиданное завершение цикла мониторинга.")
        
        except KeyboardInterrupt:
            cprint("\nМониторинг прерван пользователем. Переход к следующему шагу...", 'step')
            monitoring_interrupted = True
        
        finally:
            # В любом случае (успех, ошибка, прерывание) даем команду потоку остановиться
            if log_thread.is_alive():
                stop_event.set()
                log_thread.join() # Ждем, пока поток корректно завершится

        if monitoring_interrupted:
            # Если прервали, считаем, что пользователь хочет загрузить бэкап
            return True, backup_path, True

        # Сюда мы не должны попасть, но на всякий случай
        raise RuntimeError("Неожиданное завершение цикла мониторинга.")

    except Exception as e:
        cprint(f"\n!!! КРИТИЧЕСКАЯ ОШИБКА В ПРОЦЕССЕ ОБНОВЛЕНИЯ: {e}", 'error')
        return False, backup_path, False
    finally:
        cprint("\n  Очистка временных файлов...", 'info')
        if os.path.exists(extract_dir):
            try: shutil.rmtree(extract_dir)
            except Exception as e: print(f"  ПРЕДУПРЕЖДЕНИЕ: Не удалось удалить временную директорию '{extract_dir}': {e}")
        if os.path.exists(downloaded_archive_path):
            try: os.remove(downloaded_archive_path)
            except Exception as e: print(f"  ПРЕДУПРЕЖДЕНИЕ: Не удалось удалить скачанный архив '{downloaded_archive_path}': {e}")

def migrate_custom_libs(backup_path, source_base_dir):
    """
    Сравнивает библиотеки в старой и новой версии, и предлагает перенести недостающие.
    """
    cprint("\nШаг 3.5: Проверка на наличие кастомных библиотек...", 'step')

    # 1. Определяем пути к папкам /lib в старой (бэкап) и новой (распакованной) версиях
    old_lib_path = os.path.join(backup_path, 'exploded', 'WEB-INF', 'lib')
    new_lib_path = os.path.join(source_base_dir, 'exploded', 'WEB-INF', 'lib')

    # 2. Проверяем, существуют ли обе директории
    if not os.path.isdir(old_lib_path) or not os.path.isdir(new_lib_path):
        cprint("  Одна из папок 'lib' не найдена. Пропускаю сравнение.", 'warning')
        return

    # 3. Получаем списки файлов и находим разницу
    try:
        old_files = set(os.listdir(old_lib_path))
        new_files = set(os.listdir(new_lib_path))
        
        # Находим файлы, которые есть в старой версии, но отсутствуют в новой
        extra_files = sorted(list(old_files - new_files))

    except Exception as e:
        cprint(f"  Ошибка при сравнении содержимого папок 'lib': {e}", 'error')
        return

    # 4. Если различий нет, сообщаем и выходим
    if not extra_files:
        cprint("  Кастомные библиотеки не найдены. Состав идентичен.", 'success')
        return

    # 5. Если найдены различия, выводим их и запрашиваем действие у пользователя
    cprint("  ВНИМАНИЕ: Найдены библиотеки, отсутствующие в новой версии:", 'warning')
    for i, filename in enumerate(extra_files, 1):
        print(f"    {i}. {filename}")
    
    user_input = prompt("\nВведите номера файлов для переноса (через запятую, например: 1,3) или нажмите Enter, чтобы пропустить: ")

    if not user_input.strip():
        cprint("  Перенос кастомных библиотек пропущен пользователем.", 'info')
        return

    # 6. Парсим ввод пользователя и переносим выбранные файлы
    selected_indices = []
    parts = user_input.split(',')
    for part in parts:
        try:
            index = int(part.strip())
            if 1 <= index <= len(extra_files):
                selected_indices.append(index)
            else:
                cprint(f"  Номер {index} вне диапазона. Пропускаю.", 'warning')
        except ValueError:
            cprint(f"  Некорректный ввод '{part}'. Пропускаю.", 'warning')

    if not selected_indices:
        cprint("  Не выбрано ни одного корректного файла для переноса.", 'info')
        return

    cprint("  Перенос выбранных библиотек...", 'step')
    for index in set(selected_indices): # Используем set для исключения дубликатов
        try:
            filename_to_move = extra_files[index - 1] # -1, так как список для пользователя 1-based
            
            src_file = os.path.join(old_lib_path, filename_to_move)
            dest_file = os.path.join(new_lib_path, filename_to_move)
            
            cprint(f"    Копирование '{filename_to_move}'...", 'info')
            shutil.copy2(src_file, dest_file) # copy2 сохраняет метаданные

        except Exception as e:
            cprint(f"    !!! Ошибка при копировании файла '{filename_to_move}': {e}", 'error')
    
    cprint("  Перенос завершен.", 'success')

def upload_backup(config, source_versions_root, local_backup_path, service_name, source_type, ftp_conn=None):
    """
    Архивирует, загружает бэкап с прогресс-баром и удаляет локальную копию.
    """
    if not os.path.isdir(local_backup_path):
        cprint(f"Ошибка: Директория бэкапа '{local_backup_path}' не найдена.", 'error')
        return False

    cprint(f"\n--- Загрузка бэкапа на {source_type.upper()} ---", 'header')

    timestamp = time.strftime('%Y%m%d_%H%M%S')
    archive_base_name = f"{service_name}_backup_{timestamp}"
    temp_archive_path = os.path.join(config['General']['download_dir'], archive_base_name)
    
    cprint(f"Создание ZIP-архива '{archive_base_name}.zip'...", 'step')
    try:
        final_archive_path = shutil.make_archive(
            base_name=temp_archive_path, format='zip',
            root_dir=os.path.dirname(local_backup_path), base_dir=os.path.basename(local_backup_path)
        )
        cprint(f"Архив успешно создан: '{final_archive_path}'", 'success')
    except Exception as e:
        cprint(f"Ошибка при создании ZIP-архива: {e}", 'error')
        return False

    try:
        parent_of_root = os.path.dirname(source_versions_root)
        grandparent_of_root = os.path.dirname(parent_of_root)
        destination_dir = os.path.join(grandparent_of_root, 'temp', 'update_backups')
    except Exception as e:
        cprint(f"Не удалось рассчитать путь назначения для бэкапа: {e}", 'error')
        os.remove(final_archive_path)
        return False

    upload_successful = False
    try:
        file_size = os.path.getsize(final_archive_path)
        archive_name = os.path.basename(final_archive_path)

        if source_type == 'smb':
            cprint(f"Загрузка архива на SMB...", 'step')
            os.makedirs(destination_dir, exist_ok=True)
            destination_file_path = os.path.join(destination_dir, archive_name)
            
            # Копируем с прогресс-баром
            chunk_size = 1024 * 1024
            with open(final_archive_path, 'rb') as fsrc:
                with open(destination_file_path, 'wb') as fdst:
                    with tqdm(
                        total=file_size, unit='B', unit_scale=True, unit_divisor=1024,
                        desc=archive_name, ascii=True
                    ) as progress:
                        while True:
                            chunk = fsrc.read(chunk_size)
                            if not chunk: break
                            fdst.write(chunk)
                            progress.update(len(chunk))
            
            os.remove(final_archive_path) # Удаляем исходный файл после копирования
            cprint("Загрузка на SMB завершена.", 'success')
            upload_successful = True

        elif source_type == 'ftp':
            if not ftp_conn: raise ConnectionError("FTP соединение не установлено.")
            cprint(f"Загрузка архива на FTP...", 'step')
            
            destination_dir_ftp = destination_dir.replace('\\', '/')
            try:
                ftp_conn.cwd('/')
                path_parts = [part for part in destination_dir_ftp.split('/') if part]
                for part in path_parts:
                    try: ftp_conn.cwd(part)
                    except ftplib.error_perm:
                        ftp_conn.mkd(part)
                        ftp_conn.cwd(part)
            except ftplib.all_errors as e:
                raise IOError(f"Не удалось создать/перейти в директорию на FTP: {e}")

            # Загружаем файл с прогресс-баром
            with open(final_archive_path, 'rb') as f:
                with tqdm(
                    total=file_size, unit='B', unit_scale=True, unit_divisor=1024,
                    desc=archive_name, ascii=True
                ) as progress:
                    # Создаем коллбэк, который будет передавать РАЗМЕР блока, а не сам блок
                    def progress_callback(chunk):
                        progress.update(len(chunk))
                    
                    ftp_conn.storbinary(f'STOR {archive_name}', f, callback=progress_callback)
            
            os.remove(final_archive_path)
            cprint("Загрузка на FTP завершена.", 'success')
            upload_successful = True

    except Exception as e:
        cprint(f"Ошибка при загрузке бэкапа: {e}", 'error')
        if os.path.exists(final_archive_path):
             cprint(f"Локальный архив сохранен для ручного копирования: {final_archive_path}", 'warning')
        return False

    if upload_successful:
        cprint(f"Очистка локальной папки бэкапа: '{local_backup_path}'...", 'step')
        try:
            shutil.rmtree(local_backup_path)
            cprint("Локальный бэкап успешно удален.", 'success')
        except Exception as e:
            cprint(f"ПРЕДУПРЕЖДЕНИЕ: Не удалось удалить локальную папку бэкапа: {e}", 'warning')
    
    return upload_successful

def download_and_place_installers(server_info, update_source_info, ftp_conn=None):
    """
    Находит, загружает и размещает .exe установщики с отображением прогресс-бара.
    """
    server_type = server_info['type']
    server_dir = server_info['server_dir']
    source_type = update_source_info['source']
    source_path = update_source_info['path']

    cprint(f"  Поиск и размещение установщиков для сервера типа '{server_type}'...", 'step')

    installer_map = {
        'Setup.Front.exe': 'Front',
        'Setup.RMS.BackOffice.exe': 'BackOffice',
        'Setup.Chain.BackOffice.exe': 'BackOffice'
    }

    files_to_process = []
    if server_type == 'iikoRMS':
        files_to_process = ['Setup.Front.exe', 'Setup.RMS.BackOffice.exe']
    elif server_type == 'iikoChain':
        files_to_process = ['Setup.Chain.BackOffice.exe']
    else:
        cprint(f"  Неизвестный тип сервера '{server_type}', пропускаю загрузку.", 'warning')
        return True

    all_successful = True
    for filename in files_to_process:
        try:
            target_subdir = installer_map.get(filename)
            if not target_subdir: continue

            destination_folder = os.path.join(server_dir, 'exploded', 'update', target_subdir)
            destination_file_path = os.path.join(destination_folder, filename)
            os.makedirs(destination_folder, exist_ok=True)

            if source_type == 'smb':
                source_file_path = os.path.join(source_path, filename)
                if os.path.exists(source_file_path):
                    # --- TQDM-based copy for SMB ---
                    file_size = os.path.getsize(source_file_path)
                    chunk_size = 1024 * 1024
                    with open(source_file_path, 'rb') as fsrc:
                        with open(destination_file_path, 'wb') as fdst:
                            with tqdm(
                                total=file_size, unit='B', unit_scale=True, unit_divisor=1024,
                                desc=filename, ascii=True
                            ) as progress:
                                while True:
                                    chunk = fsrc.read(chunk_size)
                                    if not chunk: break
                                    fdst.write(chunk)
                                    progress.update(len(chunk))
                else:
                    cprint(f"    ПРЕДУПРЕЖДЕНИЕ: Установщик '{filename}' не найден на SMB.", 'warning')
                    all_successful = False

            elif source_type == 'ftp':
                if not ftp_conn: raise ConnectionError("FTP соединение не установлено.")
                
                # --- TQDM-based download for FTP ---
                absolute_ftp_path = '/' + source_path.strip('/')
                ftp_conn.cwd(absolute_ftp_path)
                
                file_list = ftp_conn.nlst()
                if filename in file_list:
                    file_size = ftp_conn.size(filename)
                    with open(destination_file_path, 'wb') as local_file:
                        with tqdm(
                            total=file_size, unit='B', unit_scale=True, unit_divisor=1024,
                            desc=filename, ascii=True
                        ) as progress:
                            def progress_callback(chunk):
                                local_file.write(chunk)
                                progress.update(len(chunk))
                            
                            ftp_conn.retrbinary(f'RETR {filename}', progress_callback)
                else:
                    cprint(f"    ПРЕДУПРЕЖДЕНИЕ: Установщик '{filename}' не найден на FTP.", 'warning')
                    all_successful = False
                
                try: ftp_conn.cwd('/')
                except: pass

        except Exception as e:
            cprint(f"!!! ОШИБКА при обработке установщика '{filename}': {e}", 'error')
            all_successful = False
            if source_type == 'ftp' and ftp_conn:
                try: ftp_conn.cwd('/')
                except: pass

    return all_successful

def monitor_log_in_thread(log_path, stop_event, success_event, success_message):
    """
    Эта функция выполняется в отдельном потоке, отслеживая лог-файл.
    :param log_path: Путь к файлу startup.log.
    :param stop_event: Событие (threading.Event) для остановки потока извне.
    :param success_event: Событие (threading.Event) для сообщения об успехе основному потоку.
    :param success_message: Строка, которую ищем в логе.
    """
    cprint("  (Поток-наблюдатель запущен...)", 'step')
    last_pos = 0
    while not stop_event.is_set():
        try:
            if not os.path.exists(log_path):
                time.sleep(1) # Ждем, если файл временно исчез
                continue

            with open(log_path, 'rb') as f:
                f.seek(last_pos)
                chunk = f.read()
                if chunk:
                    decoded_chunk = decode_bytes_with_fallbacks(chunk, fallbacks=['utf-8', 'cp1251', 'cp866'], errors='ignore')
                    if decoded_chunk:
                        lines = decoded_chunk.splitlines()
                        for line in lines:
                            print(f"    {line}")
                            if success_message in line:
                                cprint("\n  (Поток-наблюдатель нашел сообщение об успехе!)", 'success')
                                success_event.set() # Сигнализируем об успехе
                                return # Завершаем поток
                    last_pos = f.tell()
        except Exception as e:
            cprint(f"\n  (Ошибка в потоке-наблюдателе: {e})", 'warning')
        
        # Небольшая пауза, чтобы не нагружать процессор
        time.sleep(LOG_POLL_INTERVAL)
    cprint("\n  (Поток-наблюдатель остановлен)", 'step')

# --- Основная логика программы ---

if __name__ == "__main__":
    
    init() # Инициализация colorama
    cprint("--- Утилита автоматизированного обновления серверов iikoRMS ---", 'header')
    
    # Попытка повышения прав UAC
    elevate_privileges() 
    # Если мы дошли до этого места, значит, скрипт либо уже был админом, либо успешно перезапустился с UAC.
    if is_admin():
        cprint("Программа запущена от имени Администратора. Продолжаю...", 'success')
    else:
        # Этого блока по идее быть не должно, если elevate_privileges сработал
        cprint("!!! ОШИБКА: Не удалось получить права Администратора. Программа не может продолжить.", 'warning')
        prompt("Нажмите Enter для выхода.")
        sys.exit(1)
    
    # 1. Загрузка конфигурации
    config = load_config(CONFIG_FILE)
    if config is None:
        prompt("Нажмите Enter после редактирования файла config.ini...")
        config = load_config(CONFIG_FILE)
        if config is None:
             cprint("Не удалось загрузить конфигурацию после редактирования. Выход.", 'warning')
             exit()

    # Убедимся, что директории для скачивания и бэкапов существуют
    os.makedirs(config['General']['download_dir'], exist_ok=True)
    os.makedirs(config['General']['backup_dir'], exist_ok=True)

    # 2. Поиск локальных серверов
    service_keywords = config['Services'].get('keywords', DEFAULT_CONFIG['Services']['keywords'])
    found_servers = get_service_info_wmic(service_keywords) 

    if not found_servers:
        cprint("\nНе найдено ни одного подходящего сервера iikoRMS/iikoChain. Проверьте, запущены ли службы и корректны ли ключевые слова в config.ini.", 'warning')
        prompt("Нажмите Enter для выхода.")
        exit()

    # 3. Поиск доступных версий
    ftp_url = config['FTP'].get('url', '').strip()
    smb_path = config['SMB'].get('path', '').strip()

    ftp_conn = None
    ftp_base_path = ''
    available_versions = {} 

    if ftp_url:
        ftp_conn, ftp_base_path = connect_ftp(ftp_url)
        if ftp_conn: available_versions.update(find_versions_on_ftp(ftp_conn, ftp_base_path))
        else: cprint("Пропуск поиска на FTP из-за ошибки подключения.", 'warning')
    else: cprint("FTP URL не указан в config.ini. Пропуск поиска на FTP.", 'warning')

    if smb_path: available_versions.update(find_versions_on_smb(smb_path))
    else: cprint("SMB Path не указан в config.ini. Пропуск поиска на SMB.", 'warning')

    if not available_versions:
        cprint("\nНе найдено доступных версий обновлений ни на FTP, ни на SMB. Проверьте настройки в config.ini и наличие архивов в директориях.", 'error')
        if ftp_conn: 
            try: ftp_conn.quit()
            except: pass
        prompt("Нажмите Enter для выхода.")
        exit()

    # 4. Выбор сервера и версии
    cprint("\nНайденные серверы:", 'step')
    for i, server in enumerate(found_servers):
        cprint(f"{i + 1}. [{server['type']}] '{server['display_name']}' (Статус: {server['status']}, Директория: {server['server_dir']})", 'step')

    selected_server_index = -1
    while selected_server_index < 0 or selected_server_index >= len(found_servers):
        try:
            user_prompt = prompt(f"Выберите номер сервера для обновления (1-{len(found_servers)}): ")
            selected_server_index = int(user_prompt) - 1
            if selected_server_index < 0 or selected_server_index >= len(found_servers):
                 cprint("Неверный номер сервера. Попробуйте еще раз.", 'warning')
        except ValueError: cprint("Неверный ввод. Введите число.", 'error')

    selected_server = found_servers[selected_server_index]

    # --- Проверка и изменение учетной записи службы ---
    cprint(f"\n--- Проверка учетной записи службы '{selected_server['display_name']}' ---", 'header')
    current_account = get_service_logon_account_wmic(selected_server['name'])
    cprint(f"Текущая учетная запись службы: '{current_account}'", 'info')

    is_local_system = current_account and current_account.lower() in ['localsystem', 'nt authority\\system']

    if not is_local_system:
        cprint("Служба не запущена от имени 'LocalSystem'. Рекомендуется изменить учетную запись.", 'warning')
        change_account_prompt = prompt("Изменить учетную запись службы на 'LocalSystem'? (да/нет/y/n): ").lower()
        if change_account_prompt in ['да', 'y']:
            if set_service_logon_account_sc(selected_server['name'], 'LocalSystem'):
                cprint("Учетная запись службы успешно изменена на 'LocalSystem'.", 'success')
            else:
                cprint("!!! Ошибка изменения учетной записи службы. Продолжение может привести к сбоям в работе сервера.")
                proceed_anyway = prompt("Продолжить обновление, несмотря на ошибку изменения учетной записи? (да/нет/y/n): ").lower()
                if proceed_anyway not in ['да', 'y']:
                    cprint("Обновление отменено.", 'warning')
                    if ftp_conn: 
                        try: ftp_conn.quit()
                        except: pass
                    prompt("Нажмите Enter для выхода.")
                    sys.exit(1)
        else:
            cprint("Учетная запись службы не будет изменена. Продолжение может привести к сбоям в работе сервера.", 'warning')
            proceed_anyway = prompt("Продолжить обновление, несмотря на неподходящую учетную запись? (да/нет/y/n): ").lower()
            if proceed_anyway not in ['да', 'y']:
                cprint("Обновление отменено.", 'warning')
                if ftp_conn: 
                    try: ftp_conn.quit()
                    except: pass
                prompt("Нажмите Enter для выхода.")
                sys.exit(1)
    else: cprint("Учетная запись службы уже 'LocalSystem'.", 'info')
    # --- Конец проверки и изменения учетной записи ---

    cprint("\nДоступные версии для обновления:", 'step')
    version_keys = sorted(available_versions.keys()) 
    for i, key in enumerate(version_keys):
        info = available_versions[key]
        cprint(f"{i + 1}. {key} (Архивы: {', '.join(info['archives'])})", 'info') 

    selected_version_index = -1
    while selected_version_index < 0 or selected_version_index >= len(version_keys):
        try:
            user_prompt = prompt(f"Выберите номер версии для обновления (1-{len(version_keys)}): ")
            selected_version_index = int(user_prompt) - 1
            if selected_version_index < 0 or selected_version_index >= len(version_keys):
                 cprint("Неверный номер версии. Попробуйте еще раз.", 'warning')
        except ValueError: cprint("Ну ты дурак? Введи число уже. Пожалуйста", 'warning')

    selected_version_key = version_keys[selected_version_index]
    selected_version_info = available_versions[selected_version_key]

    # Автоматический выбор архива на основе типа сервера
    server_type_keyword = 'chain' if selected_server['type'] == 'iikoChain' else 'rms'
    candidate_archives = [
        arc for arc in selected_version_info['archives'] 
        if server_type_keyword in arc.lower()
    ]

    selected_archive_name = None
    if len(candidate_archives) == 1:
        selected_archive_name = candidate_archives[0]
        cprint(f"\nАвтоматически выбран архив для '{selected_server['type']}': {selected_archive_name}", 'success')
    elif len(candidate_archives) == 0:
        cprint(f"!!! ОШИБКА: Не найден подходящий архив для '{selected_server['type']}' (со словом '{server_type_keyword}') в версии {selected_version_info['version']}.", 'error')
        cprint(f"!!! Доступные архивы: {', '.join(selected_version_info['archives'])}", 'warning')
        prompt("Нажмите Enter для выхода.")
        sys.exit(1)
    else:
        cprint(f"!!! ОШИБКА: Найдено несколько подходящих архивов для '{selected_server['type']}' в версии {selected_version_info['version']}: {', '.join(candidate_archives)}", 'error')
        cprint("!!! Невозможно сделать однозначный выбор. Уточните содержимое папки с обновлениями.", 'warning')
        prompt("Нажмите Enter для выхода.")
        sys.exit(1)

    selected_version_info['selected_archive'] = selected_archive_name


    cprint(f"\n--- Подтверждение обновления ---", 'header')
    cprint(f"Сервер: [{selected_server['type']}] '{selected_server['display_name']}' (Директория: {selected_server['server_dir']})", 'step')
    cprint(f"Выбранная версия: {selected_version_info['version']}", 'step')
    cprint(f"Источник: {selected_version_info['source']}", 'step')
    cprint(f"Путь источника: {selected_version_info['path']}", 'step')
    cprint(f"Выбранный архив: {selected_version_info['selected_archive']}", 'step')


    confirm = prompt("Вы уверены, что хотите продолжить? (да/нет/y/n): ").lower()
    if confirm not in ['да', 'y']:
        cprint("Обновление отменено пользователем.", 'warning')
        if ftp_conn: 
            try: ftp_conn.quit()
            except: pass
        prompt("Нажмите Enter для выхода.")
        sys.exit(0)

    # 5. Скачивание/Копирование архива
    cprint("\n--- Скачивание/Копирование архива обновления ---", 'header')
    downloaded_archive_path = download_update_archive(selected_version_info, config['General']['download_dir'], ftp_conn=ftp_conn)

    if not downloaded_archive_path:
        cprint("Ошибка при скачивании/копировании архива обновления. Процедура прервана.", 'error')
        if ftp_conn: 
            try: ftp_conn.quit()
            except: pass
        prompt("Нажмите Enter для выхода.")
        sys.exit(1)

    # 6. Выполнение процедуры обновления
    update_successful, created_backup_path, monitoring_skipped = perform_update(
        server_info=selected_server,
        downloaded_archive_path=downloaded_archive_path,
        backup_base_dir=config['General']['backup_dir'],
        selected_version_info=selected_version_info,
        ftp_conn=ftp_conn
    )

    # 7. Загрузка бэкапа на исходный ресурс
    # *** ИЗМЕНЕНИЯ ЗДЕСЬ: Новая логика для загрузки бэкапа ***
    if update_successful and created_backup_path:
        # *** ИСПРАВЛЕННАЯ ЛОГИКА ПЕРЕДАЧИ ПАРАМЕТРОВ ***
        
        # Определяем корневой путь источника обновлений из конфига
        source_type = selected_version_info['source']
        source_versions_root = ''
        if source_type == 'smb':
            source_versions_root = smb_path  # Это путь из config.ini, например \\...\\distr\\iiko
        elif source_type == 'ftp':
            source_versions_root = ftp_base_path # Это путь, полученный при подключении к FTP

        # Проверяем, что смогли определить путь
        if not source_versions_root:
            cprint("!!! Не удалось определить корневой путь источника обновлений. Загрузка бэкапа невозможна.", 'warning')
        
        # Если пользователь пропустил мониторинг, загружаем бэкап автоматически
        elif monitoring_skipped:
            cprint("\nЗапуск автоматической загрузки бэкапа...", 'header')
            upload_backup(
                config=config,
                source_versions_root=source_versions_root,
                local_backup_path=created_backup_path,
                service_name=selected_server['name'],
                source_type=source_type,
                ftp_conn=ftp_conn
            )
        else:
            # Если пользователь следил за логом, спрашиваем как раньше
            upload_confirm = prompt(f"\nЗагрузить созданный бэкап на {source_type.upper()}? (да/нет/y/n): ").lower()
            if upload_confirm in ['да', 'y']:
                upload_backup(
                    config=config,
                    source_versions_root=source_versions_root,
                    local_backup_path=created_backup_path,
                    service_name=selected_server['name'],
                    source_type=source_type,
                    ftp_conn=ftp_conn
                )
            else:
                cprint("Загрузка бэкапа пропущена. Локальный бэкап сохранен в:", created_backup_path, 'warning')

    elif not update_successful and created_backup_path:
         cprint(f"\nЗагрузка бэкапа пропущена, т.к. процедура обновления завершилась с ошибкой.", 'error')
         cprint(f"Бэкап для ручного восстановления сохранен в: {created_backup_path}", 'warning')

    if ftp_conn:
        try: ftp_conn.quit()
        except: pass

    cprint("\n--- Программа завершила работу ---", 'header')
    prompt("Нажмите Enter для выхода.")
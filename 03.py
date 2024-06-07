import os
import subprocess
import winreg


def add_path_to_env_var(path, is_user_var):
    try:
        env_var_name = "Path"
        if is_user_var:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Environment', 0, winreg.KEY_ALL_ACCESS)
        else:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 0,
                                 winreg.KEY_ALL_ACCESS)

        try:
            current_value, _ = winreg.QueryValueEx(key, env_var_name)
            new_value = f"{current_value};{path}" if path not in current_value else current_value
        except FileNotFoundError:
            new_value = path

        winreg.SetValueEx(key, env_var_name, 0, winreg.REG_EXPAND_SZ, new_value)
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False


def clear_user_var(env_var_name):
    try:
        subprocess.run(['setx', env_var_name, ''], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        return False


def delete_env_var(env_var_name, is_user_var):
    try:
        if is_user_var:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Environment', 0, winreg.KEY_SET_VALUE)
        else:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 0,
                                 winreg.KEY_SET_VALUE)

        winreg.DeleteValue(key, env_var_name)
        winreg.CloseKey(key)
        print(f"成功删除变量 '{env_var_name}'")
        return True
    except FileNotFoundError:
        print(f"变量 '{env_var_name}' 不存在")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def rename_env_var(env_var_name, new_name, is_user_var):
    try:
        if is_user_var:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Environment', 0, winreg.KEY_ALL_ACCESS)
        else:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 0,
                                 winreg.KEY_ALL_ACCESS)

        value, type_ = winreg.QueryValueEx(key, env_var_name)
        winreg.SetValueEx(key, new_name, 0, type_, value)
        winreg.DeleteValue(key, env_var_name)
        winreg.CloseKey(key)
        print(f"变量 '{env_var_name}' 重命名为 '{new_name}'")
        return True
    except FileNotFoundError:
        print(f"变量 '{env_var_name}' 不存在")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def update_env_var_value(env_var_name, new_value, is_user_var):
    try:
        if is_user_var:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Environment', 0, winreg.KEY_SET_VALUE)
        else:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 0,
                                 winreg.KEY_SET_VALUE)

        winreg.SetValueEx(key, env_var_name, 0, winreg.REG_EXPAND_SZ, new_value)
        winreg.CloseKey(key)
        print(f"变量 '{env_var_name}' 的值已更新为 '{new_value}'")
        return True
    except FileNotFoundError:
        print(f"变量 '{env_var_name}' 不存在")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False


def search_env_var(env_var_name):
    results = []

    def search_in_registry(hkey, sub_key, env_var_name):
        try:
            key = winreg.OpenKey(hkey, sub_key, 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, env_var_name)
            winreg.CloseKey(key)
            return value
        except FileNotFoundError:
            return None

    user_value = search_in_registry(winreg.HKEY_CURRENT_USER, r'Environment', env_var_name)
    if user_value is not None:
        results.append((env_var_name, user_value, '用户变量', True))

    system_value = search_in_registry(winreg.HKEY_LOCAL_MACHINE,
                                      r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', env_var_name)
    if system_value is not None:
        results.append((env_var_name, system_value, '系统变量', False))

    if not results:
        print(f"变量 '{env_var_name}' 不存在。")
    else:
        for idx, (name, value, var_type, _) in enumerate(results, start=1):
            print(f"{idx}. 变量 '{name}' 存在于 {var_type} 中，值为 '{value}'")

    return results


def check_env_var_paths():
    def check_path(path):
        return os.path.exists(path)

    def check_var_paths(hkey, sub_key, var_type, is_user_var):
        invalid_paths = []
        with winreg.OpenKey(hkey, sub_key) as key:
            for i in range(winreg.QueryInfoKey(key)[1]):
                name, value, _ = winreg.EnumValue(key, i)
                if name.lower() == "path":
                    for p in value.split(os.pathsep):
                        if "%SystemRoot%" in p:
                            continue
                        if not check_path(p):
                            invalid_paths.append((name, p, var_type, is_user_var))
        return invalid_paths

    user_invalid_paths = check_var_paths(winreg.HKEY_CURRENT_USER, r'Environment', '用户变量', True)
    system_invalid_paths = check_var_paths(winreg.HKEY_LOCAL_MACHINE,
                                           r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', '系统变量',
                                           False)

    if user_invalid_paths:
        print("以下用户变量的路径不存在：")
        for idx, (name, path, var_type, _) in enumerate(user_invalid_paths, start=1):
            print(f"{idx}. 变量名称：{name}，路径：{path}，类型：{var_type}")

    if system_invalid_paths:
        print("以下系统变量的路径不存在：")
        for idx, (name, path, var_type, _) in enumerate(system_invalid_paths, start=1):
            print(f"{idx}. 变量名称：{name}，路径：{path}，类型：{var_type}")

    return user_invalid_paths + system_invalid_paths


def main():
    while True:
        choice = input(
            "请输入选项（1: 添加路径到用户变量，2: 清除用户变量，3: 搜索变量，4: 检查路径是否存在，E: 退出）: ").strip().upper()

        if choice == '1':
            path = input("请输入要添加的路径: ").strip()
            is_user_var = input("添加到用户变量 (U) 还是系统变量 (S): ").strip().upper() == 'U'
            success = add_path_to_env_var(path, is_user_var)
            if success:
                print(f"路径已成功添加到 {'用户变量' if is_user_var else '系统变量'} 'Path' 中。")
            else:
                print(f"添加路径到 {'用户变量' if is_user_var else '系统变量'} 'Path' 失败。")
        elif choice == '2':
            env_var_name = input("请输入要清除的用户环境变量名称: ").strip()
            success = clear_user_var(env_var_name)
            if success:
                print(f"成功清除用户变量 '{env_var_name}'。")
            else:
                print(f"清除用户变量 '{env_var_name}' 失败。")
        elif choice == '3':
            env_var_name = input("请输入要搜索的变量名称: ").strip()
            results = search_env_var(env_var_name)
            if results:
                while True:
                    sub_choice = input("输入变量序号进行操作（例如 1），输入4继续搜索，输入5回到第一步: ").strip()
                    if sub_choice.isdigit():
                        index = int(sub_choice) - 1
                        if 0 <= index < len(results):
                            name, _, var_type, is_user_var = results[index]
                            action = input(
                                f"对变量 '{name}' 进行操作：输入1删除变量，输入2重命名变量，输入3更新变量值: ").strip()
                            if action == '1':
                                delete_env_var(name, is_user_var)
                            elif action == '2':
                                new_name = input("请输入新的变量名称: ").strip()
                                rename_env_var(name, new_name, is_user_var)
                            elif action == '3':
                                new_value = input("请输入新的变量值: ").strip()
                                update_env_var_value(name, new_value, is_user_var)
                            else:
                                print("无效选择，请重新输入。")
                        else:
                            print("无效的变量序号，请重新输入。")
                    elif sub_choice == '4':
                        break
                    elif sub_choice == '5':
                        return
                    else:
                        print("无效选择，请重新输入。")
        elif choice == '4':
            invalid_paths = check_env_var_paths()
            if invalid_paths:
                while True:
                    sub_choice = input("输入路径序号进行操作（例如 1），输入4继续检查，输入5回到第一步: ").strip()
                    if sub_choice.isdigit():
                        index = int(sub_choice) - 1
                        if 0 <= index < len(invalid_paths):
                            name, path, var_type, is_user_var = invalid_paths[index]
                            action = input(
                                f"对变量 '{name}' 进行操作：输入1删除变量，输入2重命名变量，输入3更新变量值: ").strip()
                            if action == '1':
                                delete_env_var(name, is_user_var)
                            elif action == '2':
                                new_name = input("请输入新的变量名称: ").strip()
                                rename_env_var(name, new_name, is_user_var)
                            elif action == '3':
                                new_value = input("请输入新的变量值: ").strip()
                                update_env_var_value(name, new_value, is_user_var)
                            else:
                                print("无效选择，请重新输入。")
                        else:
                            print("无效的变量序号，请重新输入。")
                    elif sub_choice == '4':
                        break
                    elif sub_choice == '5':
                        return
                    else:
                        print("无效选择，请重新输入。")
        elif choice == 'E':
            print("退出脚本。")
            return
        else:
            print("无效选择，请重新输入。")


if __name__ == "__main__":
    main()

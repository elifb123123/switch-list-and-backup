import paramiko
import time
import socket

def connect_and_run(host, username, password, command):
    try:
        print(f"Connecting to {host}...")

        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname=host, username=username, password=password, look_for_keys=False)

        shell = ssh.invoke_shell()
        time.sleep(1)

        shell.send(command + "\n")
        time.sleep(2)

        output = shell.recv(65535).decode(errors="ignore")

        ssh.close()

        print(f"\nOutput from {host}:\n{'-'*50}")
        print(output)
        print(f"{'-'*50}\n")

        # detect_device_type(output)
        filename = f"{host.replace('.', '_')}_output.txt"
        with open(filename, "w", encoding="utf-8") as f:
            f.write(output)


        return output

    except paramiko.AuthenticationException:
        raise
    except paramiko.SSHException as e:
        raise
    except socket.gaierror as e:
        raise
    except Exception as e:
        raise






import subprocess


# Dumps the Google Takeout data into the data folder

def run_shell_command(command):
    try:
        result = subprocess.run(command, shell=True, check=True)
        if result.stdout:
            print("Output:", result.stdout.decode())

        if result.stderr:
            print("Error:", result.stderr.decode())

    except subprocess.CalledProcessError as e:
        print(f"Command '{command}' failed with error: {e}")


if __name__ == "__main__":
    commands = [
        # 'gcloud storage cp -r gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/jamie@homesteadcabinet.net/*.zip data/jamie@homesteadcabinet.net/',   # Jamie
        'gcloud storage cp -r gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/jason@homesteadcabinet.net/*.zip data/jason@homesteadcabinet.net/',   # Jason
        # 'gcloud storage cp -r gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/klay@homesteadcabinet.net/*.zip data/klay@homesteadcabinet.net/',    # Klay
        # 'gcloud storage cp -r gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/mark@homesteadcabinet.net/*.zip data/mark@homesteadcabinet.net/',    # Mark
        # 'gcloud storage cp -r gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/shaela@homesteadcabinet.net/*.zip data/shaela@homesteadcabinet.net/',  # Shaela
        # 'gcloud storage cp -r gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/tex@homesteadcabinet.net/*.zip data/tex@homesteadcabinet.net/',       # Tex
    ]

    for command in commands:
        run_shell_command(command)

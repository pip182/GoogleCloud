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
        # 'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/jamie@homesteadcabinet.net/*.zip data/jamie@homesteadcabinet.net/',   # Jamie
        # 'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/jason@homesteadcabinet.net/*.zip data/jason@homesteadcabinet.net/',   # Jason
        # 'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/klay@homesteadcabinet.net/*.zip data/klay@homesteadcabinet.net/',    # Klay
        # 'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/mark@homesteadcabinet.net/*.zip data/mark@homesteadcabinet.net/',    # Mark
        # 'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/shaela@homesteadcabinet.net/*.zip data/shaela@homesteadcabinet.net/',  # Shaela
        # 'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/tex@homesteadcabinet.net/*.zip data/tex@homesteadcabinet.net/',       # Tex
        'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/alex@homesteadcabinet.net/*.zip data/alex@homesteadcabinet.net/',  # Alex
        'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/bart@homesteadcabinet.net/*.zip data/bart@homesteadcabinet.net/',  # Bart
        'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/bill@homesteadcabinet.net/*.zip data/bill@homesteadcabinet.net/',  # Bill
        'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/bjardine@homesteadcabinet.net/*.zip data/bjardine@homesteadcabinet.net/',  # Bjardine
        'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/cameron@homesteadcabinet.net/*.zip data/cameron@homesteadcabinet.net/',  # Cameron
        'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/derek@homesteadcabinet.net/*.zip data/derek@homesteadcabinet.net/',  # Derek
        'gcloud storage cp gs://takeout-export-7e419057-182d-41f3-b198-59966856c80b/20241203T172855Z/nick@homesteadcabinet.net/*.zip data/nick@homesteadcabinet.net/',  # Nick
    ]

    for command in commands:
        run_shell_command(command)

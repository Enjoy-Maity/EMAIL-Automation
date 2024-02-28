import subprocess

# Replace 'my_application.exe' with the path to the application you want to run
application_path = r'C:\Users\emaienj\OneDrive - Ericsson\MPBN Planning Setup Exes\MPBN Planning Task Setup - 2.2.2.exe'

# Start the application with administrator privileges and capture the output
process = subprocess.Popen(application_path, shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

# Read the output of the application
output, error = process.communicate()

# Print the output of the application
print(output.decode())
print(error.decode())
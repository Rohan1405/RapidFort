# Use Windows Server Core as a base image
FROM mcr.microsoft.com/windows/servercore:ltsc2019

# Install Python and any dependencies
RUN powershell -Command \
    Invoke-WebRequest https://www.python.org/ftp/python/3.9.6/python-3.9.6-amd64.exe -OutFile python-installer.exe; \
    Start-Process python-installer.exe -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1' -NoNewWindow -Wait; \
    Remove-Item -Force python-installer.exe

# Install pywin32 for pythoncom
RUN pip install pywin32

# Set working directory
WORKDIR /app

# Copy app code to the container
COPY . /app

# Expose port 5000
EXPOSE 5000

# Set the environment variable to avoid warnings
ENV FLASK_APP=app.py
ENV FLASK_RUN_HOST=0.0.0.0

# Install Flask
RUN pip install flask

# Start the Flask application
CMD ["flask", "run", "--host=0.0.0.0", "--port=5000"]

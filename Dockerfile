# Use a lightweight Python image that includes Python 3
FROM python:3.10-slim

# Install system utilities and build tools for C++
RUN apt-get update && apt-get install -y \
    build-essential \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Install required Python packages for excel/data generation
RUN pip install --no-cache-dir pandas xlsxwriter matplotlib

# Create a symlink from 'py' to 'python3' so C++ popen("py ...") calls work unmodified
RUN ln -s /usr/bin/python3 /usr/bin/py

# Set the working directory
WORKDIR /app

# Copy all project files into the container
COPY . .

# Compile the C++ HTTP server for Linux
RUN g++ -O2 -std=c++17 -pthread server.cpp -o server

# Expose the server port
EXPOSE 8080

# Launch the server
CMD ["./server"]

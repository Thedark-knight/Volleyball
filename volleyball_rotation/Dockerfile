FROM kivy/buildozer

# Set working directory
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Run buildozer
CMD ["buildozer", "-v", "android", "debug"]

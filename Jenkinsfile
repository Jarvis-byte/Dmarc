pipeline {
  agent any
  stages {
    stage('Checkout') {
      steps {
        // Checkout the repository from GitHub
        git 'https://github.com/Jarvis-byte/Dmarc.git'
      }
    }
    stage('Run') {
      steps {
        // Change to the directory containing your Python script
        dir('https://github.com/Jarvis-byte/Dmarc.git') {
          // Run the Python script
          sh 'python main.py'
        }
      }
    }
  }
}

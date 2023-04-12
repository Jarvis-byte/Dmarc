
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
        // Change to the directory where the repository is cloned
        dir('Dmarc') {
          // Run the Python script in the background on Windows
          bat 'start /b python main.py'
        }
      }
    }
  }
}

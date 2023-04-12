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
        dir('Dmarc') {
          // Run the Python script
          sh 'python main.py'
        }
      }
    }
  }
}

 pipeline {
  agent any
  stages {
    stage('version') {
      steps {
        sh 'python3 --version'
      }
    }
    stage('hello') {
      steps {
        bat 'start /b python3 main.py'
      }
    }
  }
}

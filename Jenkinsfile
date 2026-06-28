pipeline {
    agent any

    environment {
        SNOWFLAKE_PASSWORD = credentials('snowflake-password')
    }

    stages {
        stage('Install Dependencies') {
            steps {
                bat 'pip install snowflake-connector-python openpyxl pyyaml'
            }
        }

        stage('Run SQL') {
            steps {
                bat 'python test.py --output-format both --verbose'
            }
        }

        stage('Archive Results') {
            steps {
                archiveArtifacts artifacts: 'summary_*.xlsx, summary_*.csv, logs/**', allowEmptyArchive: true
            }
        }
    }

    post {
        success { echo 'SQL execution completed successfully.' }
        failure { echo 'SQL execution failed. Check logs.' }
    }
}

pipeline {
    agent any

    environment {
        GITHUB_TOKEN = credentials('GITHUB_TOKEN') 
        OPENAI_API_KEY = credentials('OPENAI_API_KEY')
        GMAIL_APP_PASS = credentials('GMAIL_APP_PASS')
        DOCKER_USERNAME = credentials('DOCKER_USERNAME')
        DOCKER_PASSWORD = credentials('DOCKER_PASSWORD')
        
        DESTINATION_BRANCH = 'main'
    }

    stages {       

        stage('Pull and Run Docker Image') {
            steps {
                script {
                    bat """
                        docker login -u ${DOCKER_USERNAME} --password ${DOCKER_PASSWORD}
                    """
                    bat 'docker pull 10pdocker/ai-pr-bot:testing'

                    bat """
                        docker run \
                        -e GITHUB_TOKEN=${GITHUB_TOKEN} \
                        -e GITHUB_PR_SOURCE_REPO_OWNER=${env.GITHUB_PR_SOURCE_REPO_OWNER} \
                        -e OPENAI_API_KEY=${OPENAI_API_KEY} \
                        -e GIT_REPOSITORY_URL=${GIT_URL} \
                        -e GITHUB_PR_NUMBER=${GITHUB_PR_NUMBER} \
                        -e DECIDER="GitHub" \
                        10pdocker/ai-pr-bot:testing
                    """
                }
            }
        }

        
    }   
}

// @Library(['ter-d4t-sharedlib']) _

def config
def pipelineUtils

pipeline {
    agent none

    environment
    {
        CONFIG_DIR = ".devops"
        CONFIG_FILE = "${CONFIG_DIR}/config.json"
        SOLUTION = "Common.sln"
        TEST_PROJECT = "CommonLib.Test/CommonLib.Test.csproj"
        TOOLS_DIR = ".devops/dotnet-tools"
        PATH = "/usr/local/share/dotnet:/opt/homebrew/bin:/Users/neko0824/.dotnet/tools:/usr/bin:/bin:/usr/sbin:/sbin:${env.PATH}"
    }

    stages {
        stage('Load D4T Configuration') {
            agent { label 'service-node' }

            steps {
                script {
                    pipelineUtils = load '.devops/PipelineUtils.groovy'
                    // config = loadConfiguration(configFile: "${CONFIG_FILE}")
                     config = readJSON file: CONFIG_FILE
                }
            }
        }

        // stage('Setup Agents') {
        //     options {
        //         timeout(time: 45, unit: 'MINUTES')
        //     }

        //     agent { label 'service-node' }

        //     steps {
        //         script {
        //             config = deployAgents(config)
        //         }
        //     }
        // }

        stage('Build') {
            agent { label 'swarm' }

            stages {
                stage('Clean and Restore') {
                    steps {
                        echo 'Cleaning and Restoring NuGet Packages...'
                        sh "dotnet clean ${env.SOLUTION}"
                        sh "dotnet restore ${env.SOLUTION}"
                    }
                }

                stage('Compile') {
                    steps {
                        echo 'Building Common.sln...'
                        sh "dotnet build ${env.SOLUTION} --configuration Release"
                    }
                }

                stage('Lint') {
                    options {
                        timeout(time: 5, unit: 'MINUTES')
                    }
                    steps {
                        sh "dotnet format ${env.SOLUTION} --verify-no-changes --exclude-diagnostics CA1502 CA1505"
                    }
                }

                stage('Install DotNet Tools') {
                    steps {
                        sh "dotnet tool update --tool-path ${env.TOOLS_DIR} dotnet-reportgenerator-globaltool --version 5.3.11"
                    }
                }

                stage('Metrics') {
                    steps {
                        dir('.devops') {
                            sh('python3 -m pip install pip_system_certs lxml tabulate pyyaml')
                            sh('python3 metrics_calculate.py ./metrics_config.yaml ./metrics_reports')
                        }
                    }
                    post {
                        always {
                            archiveArtifacts artifacts: '**/metrics_reports/**', caseSensitive: true
                        }
                    }
                }

                stage('Unit Test') {
                    steps {
                        echo 'Running Unit Tests with Code Coverage...'
                        sh "dotnet test ${env.TEST_PROJECT} --configuration Release --no-build --collect:\"XPlat Code Coverage\" --results-directory .devops/TestResults"
                    }
                }

                stage('Coverage') {
                    steps {
                        sh "${env.TOOLS_DIR}/reportgenerator -reports:.devops/TestResults/**/coverage.cobertura.xml -targetdir:.devops/coverage_reports \"-reporttypes:Html;Cobertura\""
                    }
                    post {
                        always {
                            archiveArtifacts artifacts: '.devops/coverage_reports/**, .devops/TestResults/**', allowEmptyArchive: true
                            script {
                                pipelineUtils?.publishCoverageStatus()
                            }
                        }
                    }
                }
            }
        }
    }

    post {
        always {
            script {
                node('service-node') {
                    // releaseAgent()

                    try {
                        // pipelineUtils?.sendBuildEmail(config)
                    } finally {
                        cleanWs()
                    }
                }
            }
        }
    }
}

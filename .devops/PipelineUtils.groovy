// Loaded via `load()` from the Jenkinsfile so the post-build helpers below aren't inlined into
// the pipeline script itself. This repo has no shared library of its own -- it's a plain Groovy
// file evaluated in the calling script's CPS context, so it can call the same steps (bat,
// withCredentials, emailext, echo) the Jenkinsfile already does, with no new sandbox approval.

def sendBuildEmail(config) {
    def emailBody = '$DEFAULT_CONTENT <br /> <br /> <div style="padding-left: 30px; padding-bottom: 15px;"> ${CHANGES, showPaths=true, format="<div><b>%a</b>: %r %p </div><div style=\\"padding-left:30px;\\"> &#8212; &#8220;<em>%m</em>&#8221;</div>", pathFormat="</div><div style=\\"padding-left:30px;\\">%p"} </div>'

    def emailArgs = [
        subject: '$DEFAULT_SUBJECT',
        body: emailBody,
        mimeType: 'text/html',
        attachLog: true,
        from: 'tagswbuild@teradyne.com',
        // recipientProviders: [developers(), culprits(), requestor()]
    ]

    def emails = config?.Settings?.EMAILS
    if (emails) {
        emailArgs.to = emails
    }

    emailext(emailArgs)
}

// Neither githubNotify nor publishChecks is usable on this Jenkins instance (no GitHub plugin,
// no GitHub App configured for Checks). Reuse the GitHub-Account credential the checkout stages
// already authenticate with instead -- a plain PAT (GIT_ASKPASS-based, not SSH) that also works
// against the classic Statuses API. Never fails the build over a PR-visible nicety: catches
// Throwable, not Exception, since a missing/mistyped credential can surface as an Error (e.g.
// NoSuchMethodError did when githubNotify turned out not to exist here).
def publishCoverageStatus() {
    try {
        withCredentials([usernamePassword(credentialsId: 'GitHub-Account', usernameVariable: 'GH_USER', passwordVariable: 'GITHUB_STATUS_TOKEN')]) {
            bat '@py .devops\\coverage_post_github_status.py TestResults'
        }
    } catch (Throwable t) {
        echo "WARNING: could not publish coverage status to GitHub: ${t}"
    }
}

// Mutation testing is expensive (hours, not minutes), so it's gated to the branches where it's
// worth the cost: main itself, or a branch already attached to an open pull request. env.CHANGE_ID
// is only set when Jenkins is building a PR-specific ref (not just any branch that happens to have
// one open) -- this instance's branch source has no trait distinguishing draft PRs from ready ones,
// so "has an open PR" is the closest available signal for "ready" rather than a true draft check.
def shouldRunOnMainOrReadyMr() {
    if (env.BRANCH_NAME == 'main') {
        return true
    }
    return env.CHANGE_ID != null
}

return this

# Contributing to MiniExcel

Thank you for your enthusiasm and generosity in wanting to contribute to the project!

We kindly ask you to follow the guidelines hereby specified. 

## Reporting an issue

Bugs are tracked via [GitHub issues](https://github.com/mini-software/miniexcel/issues/).
The issue template will help you on the way to create a new one:

* Create an issue via the [issues list](https://github.com/mini-software/miniexcel/issues/new?template=bug_report.md)
* Include the versions of MiniExcel that are affected
* Specify the target framework and operating system
* Describe the differences between the expected behaviour and the actual behaviour
* Provide a minimum code sample that reproduces the issue

## Requesting a feature/enhancement

Similarly, feature requests are also tracked via [GitHub](https://github.com/mini-software/miniexcel/issues/) issues.
The issue template will help you on the way once again:

* Create an issue via the [issues list](https://github.com/mini-software/MiniExcel/issues/new?template=feature_request.md) 
* Explain your motivations for the feature to be included
* Describe how the feature would work
* Include workarounds where applicable

## Making a PR

**Please do not open a PR that does not reference an issue**.
* If an issue does not already exist please create one via the issues list 
* Fork the repository and create a branch with a descriptive name
* Attempt to make commits of logical units
* Verify the solution builds successfully and all tests pass
* Create the PR making sure:
    * It references the issue addressed by the PR
    * Unit tests for the changes have been added
    * Documentation has been updated where necessary

### Automatic review requests

PRs targeting `master` automatically request active members of the
[`core-reviewers` team](https://github.com/orgs/mini-software/teams/core-reviewers)
when opened, reopened, marked ready for review, or updated. The PR author,
already requested reviewers, and anyone who has already submitted a review are
excluded. Draft PRs do not request reviews. Merge approval requirements are
configured separately from review invitations.

Maintainers can update the reviewer list by adding or removing team members;
pending organization invitations must be accepted before members are included.
Removing a member does not revoke existing review requests or their separately
granted repository permissions.

Before enabling the [reviewer workflow](.github/workflows/auto-reviewers.yml),
configure a fine-grained personal access token with `mini-software` as the resource
owner and only the organization **Members: read** permission. Store it as the
repository Actions secret `CORE_REVIEWERS_TOKEN` and ensure any required
organization approval is complete. Rotate the token before it expires. The token
only reads team membership; review requests use the workflow's `GITHUB_TOKEN`
with **Pull requests: write** permission.

After the workflow is merged into `master`, use its **Run workflow** action with
a PR number to refresh reviewers on an existing PR. Team changes are otherwise
picked up on the next supported PR event, not immediately when membership changes.
The workflow uses `pull_request_target` to support fork PRs without checking out
or executing PR code. Missing credentials, team access errors, and unconfirmed
review requests fail the workflow explicitly.

### Approval requirements

The review policy for PRs targeting `master` is:

* Ordinary changes: at least **one approval** from someone other than the author.
* Important changes with the `major-change` label: at least **three approvals**.

Approvals must come from users with Write or higher repository permissions.
Invitations, comments, author approvals, bot reviews, and dismissed approvals do
not count. A later approval or change request supersedes that user's earlier
decision; a comment does not cancel an approval. Existing native protections
still apply, including blocking unresolved change requests.

Maintainers should add `major-change` before review for public API changes,
architecture changes, breaking changes, or other high-impact modifications.
Do not remove the label merely to bypass the three-approval requirement.
This is a project policy, not an ASF-wide requirement for every GitHub PR.

The [approval workflow](.github/workflows/review-approval-policy.yml) publishes
the commit status **Review approval policy**. It refreshes all open, non-draft
PRs on PR/label changes and after the
[read-only review notification workflow](.github/workflows/review-changed.yml)
completes. This two-workflow design supports fork reviews without granting the
notification workflow secrets or write permissions, or running PR code in the
privileged workflow. Neither workflow requires `CORE_REVIEWERS_TOKEN`.

Changes are processed asynchronously. Wait for the approval workflow to finish
after changing labels or reviews, then check the status before merging.
Maintainers can manually run **Review approval policy** to refresh existing PRs.
PRs sharing a head commit also share its status, so all such open, non-draft PRs
must meet their individual approval requirements.

To activate enforcement without a gap:

1. Merge these workflows into `master` while retaining the current three-approval
   native rule.
2. Run **Review approval policy** and verify its commit status on open PRs.
3. In the `main` ruleset, require the **Review approval policy** status check from
   **GitHub Actions**, then lower the native required approvals to **one** in the
   same ruleset update. Preserve the existing bypass policy.

Do not lower the native approval count before the required status check is
configured. Team-based automatic invitations additionally require the secret
described above, but tiered approval enforcement does not.

Run the reviewer and approval workflow tests with:

```sh
node --test .github/tests/auto-reviewers.test.cjs .github/tests/review-approval-policy.test.cjs
```

## AI Usage

The use of AI agents is allowed for generating bugfixes, unit tests and documentation.
Please make sure to review all AI contributions submitted on your behalf.

We kindly ask you to refrain from using AI to generate new features, 
as it becomes increasingly difficult and time-consuming for the maintainers to review them, integrate them appropriately into the codebase, and ultimately maintain them indefinitely.

## Questions?

Feel free to open a discussion or contact the maintainers directly via email if you have further inquiries.

Finally, when contributing please keep in mind our [Code of Conduct](CODE_OF_CONDUCT.md).

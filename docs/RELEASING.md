# Releasing ExcelWizard

ExcelWizard uses NuGet.org Trusted Publishing. No long-lived NuGet API key is stored in GitHub.

## Publication trigger

The `Publish NuGet package` workflow runs only when a pull request from this repository's `develop` branch is merged into `main`. Opening or closing an unmerged pull request, merging another branch, or pushing directly to `main` does not publish a package.

The workflow restores, builds, and tests the complete solution before packing and publishing `ExcelWizard`. The project version in `ExcelWizard/ExcelWizard.csproj` must be incremented before the release pull request is merged; NuGet.org rejects a version that already exists.

## GitHub configuration

The workflow uses the `nuget` GitHub environment. Its deployment branch policy allows only `main`, and its `NUGET_USER` environment secret contains the NuGet.org profile name used during the OIDC exchange.

The publish job has only `contents: read` and `id-token: write` permissions. All actions are pinned to immutable commit hashes.

## NuGet.org trusted publishing policy

Configure the policy at [NuGet.org Trusted Publishing](https://www.nuget.org/account/trustedpublishing?fromApiKeys=true) with these values:

| Field | Value |
| --- | --- |
| Repository owner | `farshaddavoudi` |
| Repository | `ExcelWizard` |
| Workflow file | `publish-nuget.yml` |
| Environment | `nuget` |

Select the owner that owns the `ExcelWizard` package and grant the policy permission to publish new versions of that package. NuGet.org expects only the workflow filename, not the `.github/workflows/` path.

The first successful publication activates a pending policy permanently when NuGet.org needs to resolve the GitHub repository and owner IDs. A pending policy is temporarily active for seven days and can be restarted from NuGet.org if it expires before the first publication.

For more information, see Microsoft's [Trusted Publishing documentation](https://learn.microsoft.com/nuget/nuget-org/trusted-publishing).

# IssueFlow production application — P4.8

This is the single application source. See [the repository README](../README.md) for setup and deployment.

From this folder:

```powershell
npm.cmd ci
npm.cmd run lint
npm.cmd test
npm.cmd run build
```

Use `manifests/issueflow-production-dev-p4-8.xml` after deploying this build to its configured HTTPS host. [Release notes and acceptance checks](docs/P4_8_RELEASE.md).

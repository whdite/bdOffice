# Building BigDogOffice

## Prerequisites

- Node.js 24.14 or later in the 24.x line, npm 10 or later.
- Stable Rust with edition 2024 support, plus a native linker/compiler.
- Windows builds use Visual Studio Build Tools with C++ and Windows SDK. The OCR helper in `packages/pdf2docx/ocr-helper` uses the Windows .NET Framework compiler and Windows SDK metadata.
- macOS: Xcode command-line tools. Packaging the Vision OCR helper requires Swift.
- Linux: a C/C++ toolchain and the desktop libraries required by Electron on the target distribution.

Install dependencies from the public registry with `npm ci`. No private npm registry or maintainer login is needed. Dependencies and downloaded native SDK packages retain their original licenses.

## Desktop

```sh
npm run typecheck
npm run build:all
npm exec -- electron apps/shell
```

`build:all` builds all six editor renderers, the Rust XLSX sidecar, the Office CLI and the Shell main and preload bundles. Rebuild the Shell after changing its main process or preload.

After `npm ci`, `npm run dev` starts the desktop Shell and the renderer development servers. It compiles the Shell and any stale editor preloads locally. To open spreadsheets in a fresh checkout, build the XLSX engine once with `npm run native:build -w @genoffice/sheets`. Development runs use a separate `GenOffice Dev` profile by default. Workspace packages resolve locally through npm workspaces.

## Packages

```sh
npm run dist:win
npm run dist:mac
npm run dist:linux
```

Run the matching platform command on a host with its native prerequisites. The default local build has no maintainer signing credentials. To distribute your own signed binaries, configure electron-builder's normal signing environment and your own update endpoint. Do not reuse another publisher's identity.

`npm run notices` generates third-party notices. `npm run licenses` checks the locked production dependencies and documented component exceptions. Keep the lockfile and license/source notices when changing dependencies.

## Android and Remote

```sh
npm run build:h5 -w @bigdogoffice/mobile
npm run mobile:android
npm run remote:package
```

The Android native build additionally requires JDK 17 (`JAVA_HOME`), Android SDK 35 (`ANDROID_HOME`) and Gradle 8.11.1 (`BIGDOG_GRADLE_HOME`). See [MOBILE.md](MOBILE.md) for local signing and [REMOTE.md](REMOTE.md) for self-hosting. Browser H5 builds do not supply Android native RTC or PDF bridges.

Build outputs go under ignored `out`, `dist`, `target` or `.build` directories. Do not commit generated packages, local SDK paths, account settings or signing material.

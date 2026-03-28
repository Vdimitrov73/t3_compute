# Building the .exe

This document explains how to build a standalone `t3_compute.exe` from the
source code, why Windows shows a SmartScreen warning, and what that warning
actually means.

---

## Why does Windows show a SmartScreen warning?

When you download and run `t3_compute.exe`, Windows may show this:

> **"Windows protected your PC"**
> Microsoft Defender SmartScreen prevented an unrecognized app from starting.

**This does not mean the file is malicious.** It means the executable is not
signed with an EV (Extended Validation) code signing certificate. EV
certificates cost approximately $300–500 USD per year and are issued by
certificate authorities like DigiCert or Sectigo after verifying the
publisher's identity. Without one, SmartScreen flags any new executable
until it accumulates enough download history to build a reputation.

To avoid this warning entirely, install from the **Microsoft Store** instead:
the Store version is signed by Microsoft and installs without any warnings.
Available in the Canadian Microsoft Store — search for **T3 Compute**.

If you prefer the direct `.exe` download, bypass the warning like this:
1. Click **"More info"**
2. Click **"Run anyway"**

If you prefer not to trust the pre-built `.exe`, you can build it yourself
from the source code in under five minutes — instructions below. You will
get the exact same result, but built on your own machine, which Windows
trusts automatically.

---

## Building the .exe yourself

### Prerequisites

- Python 3.9 or later installed and on your PATH
- The t3_compute source code (clone or download the ZIP from GitLab)

### Step 1 — Install PyInstaller and dependencies

```
pip install pyinstaller pdfplumber openpyxl xlrd
```

### Step 2 — Build the executable

From the `t3_compute` folder:

```
pyinstaller --onefile --console --name t3_compute run_t3.py
```

This produces `dist\t3_compute.exe`. The build takes 30–60 seconds.

### Step 3 — Run it

```
dist\t3_compute.exe
```

Or copy `t3_compute.exe` anywhere you like — it is fully self-contained.

---

## What PyInstaller bundles

PyInstaller packages the Python interpreter, all imported libraries
(`pdfplumber`, `openpyxl`, `xlrd`, etc.), and your script into a single
executable. Nothing is sent anywhere — the tool reads files from your local
disk and writes results to your local disk. There is no network access,
no telemetry, and no data collection of any kind.

You can verify this by inspecting the source code on GitLab before building.

---

## Notes for the CI/CD release build (maintainer only)

The pre-built `.exe` in GitLab Releases is produced by the GitLab CI
workflow `.gitlab-ci.yml` on a clean Windows runner.
The build command is identical to Step 2 above. The resulting binary is
attached to the release automatically.

To trigger a new build, create and push a version tag:

```
git tag -a v1.5.0 -m "v1.5.0"
git push origin v1.5.0
```

---

## Verifying the pre-built .exe (optional)

The SHA-256 hash of each release binary is listed in the release notes
on GitLab. To verify on Windows:

```
certutil -hashfile t3_compute.exe SHA256
```

Compare the output to the hash in the release notes. If they match, the
file is byte-for-byte identical to what was built by the CI workflow.

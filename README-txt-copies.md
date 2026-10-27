# Temporary VBA text attachments

Copy the included `tools/scripts/` files into the same location in your repository.
Use the script for your operating system; neither needs Python, Git, packages,
or access to the Internet. Bash is included with macOS; Windows PowerShell 5.1
is included with Windows 10/11. There is no need to install PowerShell on the Mac.

From your repository root, run:

**macOS Terminal**

```sh
bash ./tools/scripts/make-txt-copies.sh
```

**Windows PowerShell or Command Prompt**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\scripts\win\make-txt-copies.ps1
```

The Windows command applies its execution-policy setting to that process only;
it does not change your saved policy. Organization-enforced policy may still
prevent script execution.

Each script determines the repository root from its own location, so it can be invoked from any working directory. Keep the Bash script in tools/scripts/ and the PowerShell script in tools/scripts/win/. The Bash script resolves two parent directories; the PowerShell script resolves three. Neither requires Git to be installed.

Each run removes the existing `txt-copies/` tree, including hidden files and
subdirectories, then recreates it with the current `.bas` and `.cls` copies.
Do not store anything you need to keep in that directory.

| Source | Generated attachment |
| --- | --- |
| `cLog.cls` | `txt-copies/cLog_cls.txt` |
| `src/cTable.cls` | `txt-copies/src/cTable_cls.txt` |
| `src/helpers/JsonConverter.bas` | `txt-copies/src/helpers/JsonConverter_bas.txt` |

Extensions are matched without regard to case; output suffixes are always
`_bas.txt` or `_cls.txt`. Copies preserve the original bytes, including encoding,
line endings, VBA headers, and attributes. Only folders needed for copies are
created. The original files are never edited.

Traversal skips the root `txt-copies/` directory, `.git` directories, symbolic
links, and Windows junctions/reparse points. Otherwise all matching files are
included, including hidden and Git-ignored source folders. A nested directory
named `txt-copies` outside the root output directory is treated as ordinary source.
An existing output path that is a link or a regular file is rejected. The Windows
script also rejects links/junctions inside an old output tree before cleanup.

The scripts report a nonzero exit status on errors. A copy failure can leave a
partial output tree; correct the error and rerun. Do not run concurrent instances.

Keep your existing `txt-copies/` entry in `.gitignore`. The scripts themselves
can be version controlled; generated attachments stay temporary.

Validation: the Bash script was exercised on Linux with nested folders, spaces,
mixed-case extensions, binary/CRLF content, stale and hidden output, symlinks,
repeat runs, no matching sources, and execution outside the repo. Windows
PowerShell and macOS execution were not available for runtime validation.

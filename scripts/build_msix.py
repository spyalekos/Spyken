"""Build a self-contained Microsoft Store MSIX package (and optional local USB test kit) for Spyken."""

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from PIL import Image

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import flet_desktop.version  # noqa: E402


def read_app_version() -> str:
    main_text = (REPO_ROOT / "main.py").read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', main_text, flags=re.MULTILINE)
    if not m:
        raise RuntimeError("Could not locate APP_VERSION in main.py")
    return m.group(1)


APP_VERSION = read_app_version()

# Microsoft Partner Center Product Identity
IDENTITY_NAME = "spyalekos.Spyken"
IDENTITY_PUBLISHER = "CN=50D01956-A934-4016-AE78-743A72F6B859"
PUBLISHER_DISPLAY_NAME = "spyalekos"
DISPLAY_NAME = "Spyken"
DESCRIPTION = (
    "Ψηφιακός Εκφωνητής — Μετατροπή εγγράφων (.docx, .pdf) σε φυσική ομιλία MP3, "
    "συγχρονισμένα βίντεο παρουσίασης MP4 και Marquee banner"
)
MSIX_VERSION = f"{APP_VERSION}.0"  # Store requires 4-part Major.Minor.Build.0
BG_COLOR = "#05060D"


def find_sdk_tool(tool_name: str) -> Path:
    pattern = rf"C:\Program Files (x86)\Windows Kits\10\bin\*\x64\{tool_name}"
    matches = sorted(glob.glob(pattern), reverse=True)
    if not matches:
        raise FileNotFoundError(f"{tool_name} not found in Windows 10/11 SDK.")
    return Path(matches[0])


def find_flet_client_dir() -> Path:
    flet_ver = flet_desktop.version.version
    client_root = Path.home() / ".flet" / "client"
    exact = client_root / f"flet-desktop-full-{flet_ver}" / "flet"
    if (exact / "flet.exe").is_file():
        return exact
    # Fallback to fingerprinted cache dir for the same version
    for p in sorted(client_root.glob(f"flet-desktop-full-{flet_ver}-*")):
        candidate = p / "flet"
        if (candidate / "flet.exe").is_file():
            return candidate
    raise FileNotFoundError(
        f"Cached Flet desktop client v{flet_ver} not found in {client_root}."
    )


def generate_msix_assets(icon_path: Path, splash_path: Path, assets_out: Path) -> None:
    """Generate all required Microsoft Store visual assets from icon.png and spyken_splash.jpg."""
    assets_out.mkdir(parents=True, exist_ok=True)
    base_icon = Image.open(icon_path).convert("RGBA")
    base_splash = Image.open(splash_path).convert("RGBA") if splash_path.is_file() else base_icon

    # Square icons (centered with clean padding on transparent canvas)
    square_specs = [
        ("Square44x44Logo.png", (44, 44), 1.0),
        ("Square44x44Logo.scale-200.png", (88, 88), 1.0),
        ("Square44x44Logo.targetsize-24_altform-unplated.png", (24, 24), 1.0),
        ("Square44x44Logo.targetsize-48_altform-unplated.png", (48, 48), 1.0),
        ("Square150x150Logo.png", (150, 150), 0.84),
        ("Square150x150Logo.scale-200.png", (300, 300), 0.84),
        ("StoreLogo.png", (50, 50), 0.96),
        ("StoreLogo.scale-200.png", (100, 100), 0.96),
    ]
    for filename, (w, h), scale in square_specs:
        canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        side = max(1, int(min(w, h) * scale))
        resized = base_icon.resize((side, side), Image.Resampling.LANCZOS)
        canvas.paste(resized, ((w - side) // 2, (h - side) // 2), resized)
        canvas.save(assets_out / filename, format="PNG")

    # Wide tiles (icon centered on transparent canvas; manifest sets BackgroundColor)
    wide_specs = [
        ("Wide310x150Logo.png", (310, 150), 0.82),
        ("Wide310x150Logo.scale-200.png", (620, 300), 0.82),
    ]
    for filename, (w, h), scale in wide_specs:
        canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        side = max(1, int(min(w, h) * scale))
        resized = base_icon.resize((side, side), Image.Resampling.LANCZOS)
        canvas.paste(resized, ((w - side) // 2, (h - side) // 2), resized)
        canvas.save(assets_out / filename, format="PNG")

    # Splash screens (full Spyken splash logo centered on transparent canvas with manifest bg)
    splash_specs = [
        ("SplashScreen.png", (620, 300), 0.90),
        ("SplashScreen.scale-200.png", (1240, 600), 0.90),
    ]
    for filename, (w, h), scale in splash_specs:
        canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        side = max(1, int(min(w, h) * scale))
        resized = base_splash.resize((side, side), Image.Resampling.LANCZOS)
        canvas.paste(resized, ((w - side) // 2, (h - side) // 2), resized)
        canvas.save(assets_out / filename, format="PNG")


def write_manifest(staging_dir: Path) -> None:
    manifest_xml = f"""<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
  IgnorableNamespaces="uap rescap">

  <Identity
    Name="{IDENTITY_NAME}"
    Publisher="{IDENTITY_PUBLISHER}"
    Version="{MSIX_VERSION}"
    ProcessorArchitecture="x64" />

  <Properties>
    <DisplayName>{DISPLAY_NAME}</DisplayName>
    <PublisherDisplayName>{PUBLISHER_DISPLAY_NAME}</PublisherDisplayName>
    <Logo>Assets\\StoreLogo.png</Logo>
  </Properties>

  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.17763.0" MaxVersionTested="10.0.26100.0" />
  </Dependencies>

  <Resources>
    <Resource Language="el-GR" />
    <Resource Language="en-US" />
  </Resources>

  <Applications>
    <Application
      Id="Spyken"
      Executable="Spyken.exe"
      EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements
        DisplayName="{DISPLAY_NAME}"
        Description="{DESCRIPTION}"
        BackgroundColor="{BG_COLOR}"
        Square150x150Logo="Assets\\Square150x150Logo.png"
        Square44x44Logo="Assets\\Square44x44Logo.png">
        <uap:DefaultTile Wide310x150Logo="Assets\\Wide310x150Logo.png">
          <uap:ShowNameOnTiles>
            <uap:ShowOn Tile="square150x150Logo" />
            <uap:ShowOn Tile="wide310x150Logo" />
          </uap:ShowNameOnTiles>
        </uap:DefaultTile>
        <uap:SplashScreen Image="Assets\\SplashScreen.png" BackgroundColor="{BG_COLOR}" />
      </uap:VisualElements>
    </Application>
  </Applications>

  <Capabilities>
    <Capability Name="internetClient" />
    <rescap:Capability Name="runFullTrust" />
  </Capabilities>
</Package>
"""
    (staging_dir / "AppxManifest.xml").write_text(manifest_xml, encoding="utf-8")


def build_usb_test_kit(store_msix: Path) -> None:
    """Create dist\\usb\\ with a self-signed copy of the MSIX for testing on another PC."""
    signtool = find_sdk_tool("signtool.exe")
    usb_dir = REPO_ROOT / "dist" / "usb"
    if usb_dir.exists():
        shutil.rmtree(usb_dir)
    usb_dir.mkdir(parents=True, exist_ok=True)

    test_msix_name = f"Spyken_{MSIX_VERSION}_x64_TEST.msix"
    test_msix = usb_dir / test_msix_name
    cer_name = "Spyken_test.cer"
    cer_path = usb_dir / cer_name

    # Never sign the original Store MSIX — copy it first!
    shutil.copy2(store_msix, test_msix)

    ps_script = (
        "$ErrorActionPreference = 'Stop'; "
        "Import-Module Microsoft.PowerShell.Security -ErrorAction SilentlyContinue; "
        "Import-Module PKI -ErrorAction SilentlyContinue; "
        f"$cert = New-SelfSignedCertificate -Type Custom -Subject '{IDENTITY_PUBLISHER}' "
        "-KeyUsage DigitalSignature -FriendlyName 'Spyken Local Test Certificate' "
        "-CertStoreLocation 'Cert:\\CurrentUser\\My' "
        "-TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3', '2.5.29.19={text}'); "
        "try { "
        f"Export-Certificate -Cert $cert -FilePath '{cer_path}' -Force | Out-Null; "
        f"& '{signtool}' sign /fd SHA256 /sha1 $cert.Thumbprint '{test_msix}'; "
        "if ($LASTEXITCODE -ne 0) { throw \"signtool failed with exit code $LASTEXITCODE\" }; "
        "Write-Output \"THUMBPRINT=$($cert.Thumbprint)\" "
        "} finally { "
        "Remove-Item \"Cert:\\CurrentUser\\My\\$($cert.Thumbprint)\" -Force -ErrorAction SilentlyContinue "
        "}"
    )
    res = subprocess.run(
        ["pwsh", "-NoProfile", "-Command", ps_script]
        if shutil.which("pwsh")
        else ["powershell", "-NoProfile", "-Command", ps_script],
        capture_output=True,
        text=True,
    )
    if res.returncode != 0:
        raise RuntimeError(
            f"Self-signed USB test kit signing failed (exit {res.returncode}):\nSTDOUT:\n{res.stdout}\nSTDERR:\n{res.stderr}"
        )
    thumbprint = ""
    for line in res.stdout.splitlines():
        if line.startswith("THUMBPRINT="):
            thumbprint = line.split("=", 1)[1].strip()
    if not thumbprint:
        raise RuntimeError(f"Could not extract test certificate thumbprint.\n{res.stdout}\n{res.stderr}")

    install_bat = f"""@echo off
chcp 65001 >nul
setlocal

set "DIR=%~dp0"
set "CER=%DIR%{cer_name}"
set "MSIX=%DIR%{test_msix_name}"

echo ============================================================
echo   Spyken - Dokimastiki Egkatastasi MSIX (Test Build)
echo ============================================================
echo.

echo [1/2] Egkatastasi dokimastikou pistopoiitikou (tha zitithei UAC)...
powershell -NoProfile -Command "$arg = '-addstore -f Root ' + [char]34 + $env:CER + [char]34; $p = Start-Process -FilePath 'certutil.exe' -ArgumentList $arg -Verb RunAs -Wait -PassThru; exit $p.ExitCode"
if errorlevel 1 (
    echo [SFALMA] Den egine egkatastasi tou pistopoiitikou.
    pause
    exit /b 1
)

echo.
echo [2/2] Egkatastasi paketou MSIX gia ton trechonta christi...
powershell -NoProfile -Command "try {{ Add-AppxPackage -Path $env:MSIX -ErrorAction Stop; Write-Host 'EPITYCHIA! To Spyken egkatastathike. Anoixte to apo to menou Enarxi.' -ForegroundColor Green }} catch {{ Write-Host $_.Exception.Message -ForegroundColor Red; Write-Host 'Anoigma me to Windows App Installer...'; Start-Process -FilePath $env:MSIX }}"

echo.
pause
"""
    (usb_dir / "1_Install_Spyken_TEST.bat").write_text(install_bat, encoding="utf-8")

    cleanup_bat = f"""@echo off
chcp 65001 >nul
setlocal

echo ============================================================
echo   Spyken - Afairesi Dokimastikou Pistopoiitikou ^& App
echo ============================================================
echo.

echo [1/2] Apegkatastasi Spyken (an einai egkatestimeno)...
powershell -NoProfile -Command "Get-AppxPackage -Name '{IDENTITY_NAME}' | Remove-AppxPackage -ErrorAction SilentlyContinue"

echo [2/2] Afairesi dokimastikou pistopoiitikou (tha zitithei UAC)...
powershell -NoProfile -Command "Start-Process -FilePath 'certutil.exe' -ArgumentList '-delstore','Root','{thumbprint}' -Verb RunAs -Wait"

echo.
echo Oloklirothike o katharismos!
pause
"""
    (usb_dir / "2_Cleanup_Test_Cert.bat").write_text(cleanup_bat, encoding="utf-8")

    readme_usb = f"""Spyken - Οδηγίες Δοκιμαστικής Εγκατάστασης MSIX σε Άλλον Υπολογιστή
=====================================================================

Περιεχόμενα φακέλου:
- {test_msix_name} : Το δοκιμαστικό πακέτο MSIX (αυτοϋπογεγραμμένο)
- {cer_name}               : Το δημόσιο κλειδί του δοκιμαστικού πιστοποιητικού
- 1_Install_Spyken_TEST.bat     : Αυτόματη εγκατάσταση (1 κλικ)
- 2_Cleanup_Test_Cert.bat       : Καθαρή αφαίρεση εφαρμογής & δοκιμαστικού πιστοποιητικού

---------------------------------------------------------------------
ΤΡΟΠΟΣ Α - Αυτόματη Εγκατάσταση (Προτεινόμενος):
---------------------------------------------------------------------
1. Αντιγράψτε όλον τον φάκελο "usb" στον άλλον υπολογιστή (Windows 10/11 x64).
2. Κάντε διπλό κλικ στο αρχείο:  1_Install_Spyken_TEST.bat
3. Πατήστε "Ναι" στο παράθυρο UAC για να προστεθεί το δοκιμαστικό πιστοποιητικό.
4. Μετά την ολοκλήρωση, ανοίξτε το "Spyken" από το μενού Έναρξη των Windows!

---------------------------------------------------------------------
ΤΡΟΠΟΣ Β - Χειροκίνητη Εγκατάσταση:
---------------------------------------------------------------------
1. Διπλό κλικ στο "{cer_name}" -> "Εγκατάσταση πιστοποιητικού"
   -> "Τοπικός υπολογιστής (Local Machine)"
   -> "Τοποθέτηση όλων των πιστοποιητικών στον ακόλουθο χώρο αποθήκευσης"
   -> "Αξιόπιστες αρχές έκδοσης πιστοποιητικών ρίζας (Trusted Root Certification Authorities)"
   -> "Επόμενο" -> "Τέλος".
2. Διπλό κλικ στο "{test_msix_name}" -> "Εγκατάσταση".

---------------------------------------------------------------------
ΑΦΑΙΡΕΣΗ ΜΕΤΑ ΤΗ ΔΟΚΙΜΗ:
---------------------------------------------------------------------
Τρέξτε το "2_Cleanup_Test_Cert.bat" για να απεγκατασταθεί το δοκιμαστικό Spyken
και να αφαιρεθεί το προσωρινό πιστοποιητικό από τον υπολογιστή.
(Τα δεδομένα στον φάκελο Έγγραφα\\Spyken δεν διαγράφονται αυτόματα.)
"""
    (usb_dir / "README_USB.txt").write_text(readme_usb, encoding="utf-8")
    print(f"Created local USB test kit: {usb_dir} (Cert Thumbprint: {thumbprint})")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Microsoft Store MSIX package for Spyken.")
    parser.add_argument("--usb", action="store_true", help="Also generate self-signed dist\\usb\\ test kit.")
    parser.add_argument("--usb-only", action="store_true", help="Only regenerate dist\\usb\\ from existing Store MSIX.")
    args = parser.parse_args()

    msix_out = REPO_ROOT / "dist" / f"Spyken_{MSIX_VERSION}_x64.msix"
    if args.usb_only:
        if not msix_out.is_file():
            raise FileNotFoundError(f"Store MSIX not found: {msix_out}")
        build_usb_test_kit(msix_out)
        return

    makeappx = find_sdk_tool("makeappx.exe")
    makepri = find_sdk_tool("makepri.exe")
    flet_client_src = find_flet_client_dir()
    print(f"[1/6] Spyken v{APP_VERSION} -> MSIX {MSIX_VERSION}")
    print(f"[1/6] Using makeappx: {makeappx}")
    print(f"[1/6] Using makepri:  {makepri}")
    print(f"[1/6] Using Flet desktop client: {flet_client_src}")

    print("[2/6] Running PyInstaller (Spyken_msix.spec --clean)...")
    subprocess.run(
        ["uv", "run", "pyinstaller", "Spyken_msix.spec", "--clean", "-y"],
        cwd=str(REPO_ROOT),
        check=True,
    )

    onedir_out = REPO_ROOT / "dist" / "Spyken_msix"
    if not (onedir_out / "Spyken.exe").is_file():
        raise FileNotFoundError(f"Expected built exe at {onedir_out / 'Spyken.exe'}")

    staging_dir = REPO_ROOT / "build" / "msix_staging"
    if staging_dir.exists():
        shutil.rmtree(staging_dir)

    print(f"[3/6] Preparing MSIX staging folder: {staging_dir}")
    shutil.copytree(onedir_out, staging_dir)
    shutil.copytree(flet_client_src, staging_dir / "flet_client")

    print("[4/6] Generating Store visual assets and AppxManifest.xml...")
    generate_msix_assets(
        REPO_ROOT / "assets" / "icon.png",
        REPO_ROOT / "assets" / "spyken_splash.jpg",
        staging_dir / "Assets",
    )
    write_manifest(staging_dir)

    print("[5/6] Generating single unified resources.pri via makepri.exe...")
    priconfig = REPO_ROOT / "build" / "priconfig.xml"
    if priconfig.exists():
        priconfig.unlink()
    subprocess.run(
        [str(makepri), "createconfig", "/cf", str(priconfig), "/dq", "el-GR", "/o"],
        check=True,
    )
    # Strip <packaging> split-PRI rules so makepri outputs ONE single resources.pri for single .msix
    cfg_text = priconfig.read_text(encoding="utf-8")
    cfg_text = re.sub(r"<packaging>.*?</packaging>", "", cfg_text, flags=re.DOTALL)
    priconfig.write_text(cfg_text, encoding="utf-8")

    subprocess.run(
        [
            str(makepri),
            "new",
            "/pr",
            str(staging_dir),
            "/cf",
            str(priconfig),
            "/mn",
            str(staging_dir / "AppxManifest.xml"),
            "/of",
            str(staging_dir / "resources.pri"),
            "/o",
        ],
        check=True,
    )

    msix_out = REPO_ROOT / "dist" / f"Spyken_{MSIX_VERSION}_x64.msix"
    if msix_out.exists():
        msix_out.unlink()

    print(f"[6/6] Packing unsigned Microsoft Store MSIX: {msix_out}")
    subprocess.run(
        [str(makeappx), "pack", "/d", str(staging_dir), "/p", str(msix_out), "/o"],
        check=True,
    )

    size_mb = msix_out.stat().st_size / (1024 * 1024)
    print(f"\nSUCCESS! Created unsigned Store MSIX package: {msix_out} ({size_mb:.2f} MB)")

    if args.usb:
        build_usb_test_kit(msix_out)


if __name__ == "__main__":
    main()

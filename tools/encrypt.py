"""Encrypt the guide page so the published file can only be read with the password.

Usage:
  GUIDE_PASSWORD=... python3 tools/encrypt.py source.html index.html

- The plain source (source.html) must never be committed.
- AES-256-GCM, key from PBKDF2-SHA256 (600,000 rounds) of the password and the
  salt in tools/salt.txt. Keep the salt fixed so "remember on this device" keeps
  working after updates; change the password to lock out old devices.
"""
import base64, json, os, sys
from pathlib import Path
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITER = 600_000
HERE = Path(__file__).resolve().parent

def main():
    src, out = sys.argv[1], sys.argv[2]
    pw = os.environ["GUIDE_PASSWORD"].encode()
    salt = bytes.fromhex((HERE / "salt.txt").read_text().strip())
    key = PBKDF2HMAC(hashes.SHA256(), 32, salt, ITER).derive(pw)
    iv = os.urandom(12)
    ct = AESGCM(key).encrypt(iv, Path(src).read_bytes(), None)
    b64 = lambda b: base64.b64encode(b).decode()
    payload = json.dumps({"s": salt.hex(), "i": b64(iv), "n": ITER, "c": b64(ct)})
    page = (HERE / "gate.html").read_text(encoding="utf-8").replace("__PAYLOAD__", payload)
    Path(out).write_text(page, encoding="utf-8")

if __name__ == "__main__":
    main()

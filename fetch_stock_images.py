"""Download royalty-free stock images from Unsplash for the YONO SBI deck.

Uses well-known stable Unsplash photo IDs (CC0 / Unsplash licence). Saves to
./assets/. Falls back gracefully if a download fails - the build script
treats missing images as optional.
"""

import os
import sys
import urllib.request
import ssl

ASSETS = "assets"
os.makedirs(ASSETS, exist_ok=True)

# Stable Unsplash direct image URLs (verified popular photos, CC0 / Unsplash
# Licence). Using `?w=1600&q=80` keeps file size reasonable while looking
# crisp on a 1080p projector.
IMAGES = {
    "hero_phone_banking.jpg":
        "https://images.unsplash.com/photo-1563013544-824ae1b704d3?w=1600&q=80",
    "india_skyline.jpg":
        "https://images.unsplash.com/photo-1570168007204-dfb528c6958f?w=1600&q=80",
    "data_analytics.jpg":
        "https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=1600&q=80",
    "tech_circuit.jpg":
        "https://images.unsplash.com/photo-1518770660439-4636190af475?w=1600&q=80",
    "team_meeting.jpg":
        "https://images.unsplash.com/photo-1521737711867-e3b97375f902?w=1600&q=80",
    "thank_you_handshake.jpg":
        "https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=1600&q=80",
    "fintech_growth.jpg":
        "https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=1600&q=80",
    "future_ai.jpg":
        "https://images.unsplash.com/photo-1620712943543-bcc4688e7485?w=1600&q=80",
}


def fetch(name, url):
    out = os.path.join(ASSETS, name)
    if os.path.exists(out) and os.path.getsize(out) > 5_000:
        print(f"  [skip] {name} already present ({os.path.getsize(out)} B)")
        return True
    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; YONO-deck-builder)"}
        )
        with urllib.request.urlopen(req, timeout=20, context=ctx) as resp:
            data = resp.read()
        if len(data) < 5_000:
            print(f"  [warn] {name} too small ({len(data)} B); skipped")
            return False
        with open(out, "wb") as f:
            f.write(data)
        print(f"  [ok]   {name} ({len(data)} B)")
        return True
    except Exception as e:
        print(f"  [fail] {name}: {e}")
        return False


def main():
    successes = 0
    failures = 0
    print(f"Fetching {len(IMAGES)} stock images into ./{ASSETS}/ ...")
    for name, url in IMAGES.items():
        if fetch(name, url):
            successes += 1
        else:
            failures += 1
    print(f"\nDone: {successes} ok, {failures} failed.")
    return 0 if successes > 0 else 1


if __name__ == "__main__":
    sys.exit(main())

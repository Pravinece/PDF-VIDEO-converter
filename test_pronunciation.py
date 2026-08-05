"""Quick test: show before/after pronunciation fixes on the actual PPTX voiceover."""
import re

PRONUNCIATION_MAP = {
    "Pradhan Mantri": "Pruh-dhaan Mun-tree",
    "Mudra": "Mood-raa",
    "Yojna": "Yoj-naa",
    "Yojana": "Yoj-naa",
    "PMMY": "P M M Y",
    "Shishu": "Shee-shoo",
    "Kishor": "Kish-ore",
    "Tarun": "Tuh-roon",
    "PNB": "P N B",
    "MSME": "M S M E",
    "CGTMSE": "C G T M S E",
    "Udyog Aadhaar": "Ud-yog Aa-dhaar",
    "lakh": "laakh",
    "Crore": "Kuh-roar",
}


def apply_pronunciation_fixes(text):
    text = re.sub(r'^(voiceover|narration|voiceover\s*/\s*narration)\s*:\s*', '', text, flags=re.IGNORECASE)
    text = text.replace('/', ' or ')
    for term, phonetic in PRONUNCIATION_MAP.items():
        pattern = re.compile(re.escape(term), re.IGNORECASE)
        text = pattern.sub(phonetic, text)
    return text.strip()


samples = [
    "Voiceover/Narration: Welcome to the presentation on PNB Pradhan Mantri Mudra Yojna scheme.",
    "Voiceover/Narration: Shishu loans cover amounts up to fifty thousand rupees. Kishor loans range from above fifty thousand rupees up to five lakh rupees. Tarun loans cover amounts above five lakh rupees.",
    "Voiceover/Narration: Loans up to ten lakh rupees are generally collateral-free and are covered under CGTMSE.",
    "Voiceover/Narration: Udyog Aadhaar registration is mandatory under the scheme.",
    "Voiceover/Narration: Pradhan Mantri Mudra Yojana (PMMY) supports non-farm micro and small businesses.",
]

print("=" * 80)
print("PRONUNCIATION FIX DEMO")
print("=" * 80)
for s in samples:
    print(f"\nBEFORE: {s}")
    print(f"AFTER : {apply_pronunciation_fixes(s)}")
print("\n" + "=" * 80)

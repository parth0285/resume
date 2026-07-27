from experience_calculator import compute_experience

# Simulated designation_history extracted from Arindam Sarkar's CV (example)
record = {
    "source_file": "arindam_sarkar.pdf",
    "designation_history": [
        {"title": "Assistant Professor", "organization": "IIT Bhubaneswar", "date_range": "May 2010 -", "category": ""},
        {"title": "Lecturer", "organization": "Thapar University", "date_range": "2007-2010", "category": "academic"},
        {"title": "Postdoctoral Fellow", "organization": "Research Institute", "date_range": "2004-2006", "category": ""},
        {"title": "Doctoral Research Fellow", "organization": "University", "date_range": "2001-2004", "category": ""},
    ],
    "administrative_roles": [],
}

out = compute_experience(record)
print("Computed experience:")
print(out["experience"]) 
print("Flags:", out.get("flags"))

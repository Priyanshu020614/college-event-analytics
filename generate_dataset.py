"""
College Event Analytics System - Synthetic Dataset Generator
-------------------------------------------------------------
Generates realistic, linked tables:
  students.csv, events.csv, registrations.csv, attendance.csv, feedback.csv
and one merged, ML-ready table:
  master_dataset.csv   (one row per registration)

Usage:  python generate_dataset.py
Change N_STUDENTS / N_EVENTS below to scale the data up or down.
All results are reproducible (fixed SEED).
"""
import numpy as np
import pandas as pd

SEED = 42
N_STUDENTS = 6000
N_EVENTS = 400
START = pd.Timestamp("2022-07-01")
END = pd.Timestamp("2025-12-15")

rng = np.random.default_rng(SEED)

# ------------------------------------------------------------------ students
departments = ["CSE", "IT", "ECE", "EEE", "Mechanical", "Civil", "MBA", "BBA", "BCA", "Biotech"]
dept_p = [0.22, 0.14, 0.12, 0.08, 0.10, 0.07, 0.08, 0.07, 0.07, 0.05]
categories = ["Technical", "Cultural", "Sports", "Workshop", "Seminar", "Career", "Social/NSS"]

dept_pref = {  # preference weights per category for each department
    "CSE": [4, 1.5, 1.2, 3, 1.5, 2.5, 1], "IT": [4, 1.5, 1.2, 3, 1.5, 2.5, 1],
    "ECE": [3.5, 1.5, 1.3, 2.8, 1.5, 2, 1], "EEE": [3, 1.5, 1.4, 2.5, 1.5, 2, 1],
    "Mechanical": [2.5, 1.6, 2.5, 2.5, 1.3, 2, 1.2], "Civil": [2.2, 1.6, 2.5, 2.2, 1.3, 2, 1.4],
    "MBA": [1.2, 2, 1.5, 2, 2.5, 4, 2], "BBA": [1.2, 2.5, 1.8, 1.8, 2.2, 3.5, 2],
    "BCA": [3, 2, 1.5, 2.5, 1.5, 2.2, 1.2], "Biotech": [2, 1.8, 1.4, 3, 3, 2, 2],
}

student = pd.DataFrame({
    "student_id": [f"S{str(i).zfill(5)}" for i in range(1, N_STUDENTS + 1)],
    "department": rng.choice(departments, N_STUDENTS, p=dept_p),
    "year": rng.choice([1, 2, 3, 4], N_STUDENTS, p=[0.30, 0.28, 0.25, 0.17]),
    "gender": rng.choice(["Male", "Female", "Other"], N_STUDENTS, p=[0.54, 0.45, 0.01]),
    "residence": rng.choice(["Hostel", "Day Scholar"], N_STUDENTS, p=[0.45, 0.55]),
})
student["cgpa"] = np.clip(rng.normal(7.4, 1.0, N_STUDENTS), 4.5, 10).round(2)
# latent engagement: how active a student is in campus life
student["engagement"] = np.clip(rng.beta(2, 3.5, N_STUDENTS) + (student["residence"] == "Hostel") * 0.08, 0.02, 1)
student["club_member"] = (rng.random(N_STUDENTS) < (0.08 + 0.5 * student["engagement"])).astype(int)

# -------------------------------------------------------------------- events
types_by_cat = {
    "Technical": ["Hackathon", "Coding Contest", "Tech Talk", "Project Expo", "Robotics Challenge"],
    "Cultural": ["Dance Competition", "Music Night", "Drama", "Fashion Show", "Cultural Fest"],
    "Sports": ["Cricket Tournament", "Football League", "Badminton Meet", "Marathon", "Athletics Day"],
    "Workshop": ["AI/ML Workshop", "Web Dev Workshop", "Design Workshop", "Soft Skills Workshop", "Data Science Workshop"],
    "Seminar": ["Guest Lecture", "Research Seminar", "Industry Panel", "Alumni Talk"],
    "Career": ["Placement Drive", "Resume Clinic", "Mock Interview", "Career Fair", "Internship Meet"],
    "Social/NSS": ["Blood Donation Camp", "Tree Plantation", "Cleanliness Drive", "Charity Run"],
}
clubs = ["Coding Club", "Cultural Committee", "Sports Council", "NSS Unit", "Placement Cell",
         "IEEE Chapter", "Entrepreneurship Cell", "Literary Society", "Robotics Club", "Photography Club"]
venues = {"Auditorium": 600, "Seminar Hall": 200, "Open Ground": 1500, "Computer Lab": 80,
          "Classroom Block": 120, "Sports Complex": 800, "Online (Zoom)": 2000}

cat_p = [0.24, 0.18, 0.10, 0.18, 0.10, 0.12, 0.08]
ev_cat = rng.choice(categories, N_EVENTS, p=cat_p)
ev_type = [rng.choice(types_by_cat[c]) for c in ev_cat]

days_span = (END - START).days
ev_date = START + pd.to_timedelta(rng.integers(0, days_span, N_EVENTS), unit="D")

events = pd.DataFrame({
    "event_id": [f"E{str(i).zfill(4)}" for i in range(1, N_EVENTS + 1)],
    "event_name": [f"{t} {d.year}" for t, d in zip(ev_type, ev_date)],
    "category": ev_cat, "event_type": ev_type, "event_date": ev_date,
    "organizer": rng.choice(clubs, N_EVENTS),
})
events["venue"] = [rng.choice(["Computer Lab", "Online (Zoom)", "Classroom Block"]) if c == "Workshop"
                   else rng.choice(["Open Ground", "Sports Complex"]) if c == "Sports"
                   else rng.choice(["Auditorium", "Seminar Hall", "Open Ground"]) if c in ("Cultural", "Social/NSS")
                   else rng.choice(["Auditorium", "Seminar Hall", "Online (Zoom)", "Classroom Block"])
                   for c in events["category"]]
events["capacity"] = [int(venues[v] * rng.uniform(0.6, 1.0)) for v in events["venue"]]
events["capacity"] = events["capacity"].clip(upper=N_STUDENTS // 3)
events["is_free"] = (rng.random(N_EVENTS) < np.where(events["category"].isin(["Seminar", "Social/NSS", "Career"]), 0.85, 0.45)).astype(int)
events["fee"] = np.where(events["is_free"] == 1, 0, rng.choice([50, 100, 150, 200, 300, 500], N_EVENTS, p=[.25, .3, .2, .13, .08, .04]))
events["duration_hours"] = np.where(events["category"].isin(["Sports", "Cultural"]), rng.choice([4, 6, 8, 10], N_EVENTS),
                                    rng.choice([1, 2, 3, 4, 6], N_EVENTS, p=[.2, .3, .25, .15, .1]))
events["has_certificate"] = (rng.random(N_EVENTS) < np.where(events["category"].isin(["Workshop", "Technical", "Seminar"]), 0.8, 0.3)).astype(int)
events["has_prizes"] = (rng.random(N_EVENTS) < np.where(events["category"].isin(["Technical", "Cultural", "Sports"]), 0.8, 0.1)).astype(int)
events["has_food"] = (rng.random(N_EVENTS) < 0.5).astype(int)
events["promotion_days"] = rng.integers(2, 31, N_EVENTS)
events["promotion_channels"] = rng.integers(1, 6, N_EVENTS)  # posters, WhatsApp, Instagram, email, class visits
events["guest_speaker_rating"] = np.where(events["category"].isin(["Seminar", "Workshop", "Career", "Technical"]),
                                          np.clip(rng.normal(3.9, 0.7, N_EVENTS), 1, 5).round(1), np.nan)
events["event_quality"] = np.clip(rng.normal(0, 1, N_EVENTS), -2.5, 2.5)  # hidden driver of satisfaction
events["day_of_week"] = events["event_date"].dt.day_name()
events["month"] = events["event_date"].dt.month
events["is_weekend"] = events["event_date"].dt.dayofweek.isin([5, 6]).astype(int)
events["time_slot"] = rng.choice(["Morning", "Afternoon", "Evening"], N_EVENTS, p=[0.35, 0.45, 0.2])
# exam months in a typical Indian college calendar: Nov-Dec, Apr-May
events["near_exams"] = events["month"].isin([11, 12, 4, 5]).astype(int)
events["is_festival_week"] = (rng.random(N_EVENTS) < 0.12).astype(int)
events["semester"] = np.where(events["month"].isin([7, 8, 9, 10, 11, 12]), "Odd", "Even")
events["academic_year"] = np.where(events["month"] >= 7, events["event_date"].dt.year, events["event_date"].dt.year - 1)

# ------------------------------------------------------------- registrations
cat_idx = {c: i for i, c in enumerate(categories)}
S_dept = student["department"].values
S_year = student["year"].values
S_eng = student["engagement"].values
S_club = student["club_member"].values
S_ids = student["student_id"].values
pref_matrix = np.array([dept_pref[d] for d in S_dept])  # N_STUDENTS x 7

reg_rows = []
for _, ev in events.iterrows():
    ci = cat_idx[ev["category"]]
    base = pref_matrix[:, ci] / 3.0
    p = 0.02 * base * (0.4 + 2.2 * S_eng)
    p *= 1 + 0.5 * S_club
    p *= {1: 1.3, 2: 1.15, 3: 1.0, 4: 0.8}[0] if False else np.select([S_year == 1, S_year == 2, S_year == 3], [1.3, 1.15, 1.0], 0.8)
    p *= 1 + 0.08 * ev["promotion_channels"] + 0.01 * ev["promotion_days"]
    p *= 1.25 if ev["has_certificate"] else 1.0
    p *= 1.3 if ev["has_prizes"] else 1.0
    p *= 1.15 if ev["has_food"] else 1.0
    p *= 1.0 if ev["is_free"] else max(0.35, 1 - ev["fee"] / 700)
    p *= 0.75 if ev["near_exams"] else 1.0
    p *= 1.2 if ev["is_festival_week"] else 1.0
    p = np.clip(p, 0, 0.9)
    chosen = np.where(rng.random(N_STUDENTS) < p)[0]
    if len(chosen) > ev["capacity"]:
        chosen = rng.choice(chosen, ev["capacity"], replace=False)
    lead = np.clip(rng.exponential(ev["promotion_days"] / 3.0, len(chosen)).round().astype(int) + 1, 1, ev["promotion_days"])
    for idx, ld in zip(chosen, lead):
        reg_rows.append((S_ids[idx], ev["event_id"], ev["event_date"] - pd.Timedelta(days=int(ld)), int(ld)))

reg = pd.DataFrame(reg_rows, columns=["student_id", "event_id", "registration_date", "days_before_event"])
reg.insert(0, "registration_id", [f"R{str(i).zfill(6)}" for i in range(1, len(reg) + 1)])
reg["registration_mode"] = rng.choice(["Online Form", "QR Code", "Spot Registration", "Club Referral"], len(reg), p=[.55, .2, .1, .15])
reg["team_size"] = np.where(reg.merge(events[["event_id", "event_type"]], on="event_id")["event_type"]
                            .isin(["Hackathon", "Robotics Challenge", "Project Expo", "Cricket Tournament", "Football League"]).values,
                            rng.integers(2, 6, len(reg)), 1)

# ---------------------------------------------------------------- attendance
m = reg.merge(events, on="event_id").merge(student, on="student_id")
logit = (0.9
         + 1.2 * (m["engagement"] - 0.3)
         + 0.7 * m["club_member"]
         + 0.55 * (m["fee"] > 0)
         + 0.25 * (m["has_food"] == 1)
         + 0.2 * (m["has_certificate"] == 1)
         - 0.6 * m["near_exams"]
         - 0.35 * (m["event_type"].isin(["Marathon"]))
         - 0.25 * (m["time_slot"] == "Morning") * (m["residence"] == "Day Scholar")
         - 0.04 * m["days_before_event"]
         + 0.25 * (m["year"] == 1)
         - 0.3 * (m["year"] == 4)
         - 0.3 * (m["venue"] == "Online (Zoom)")
         + 0.15 * m["is_weekend"]
         + rng.normal(0, 0.5, len(m)))
p_att = 1 / (1 + np.exp(-logit))
attended = (rng.random(len(m)) < p_att).astype(int)
reg["attended"] = attended
reg["attendance_probability_true"] = p_att.round(3)

att = reg.loc[reg["attended"] == 1, ["registration_id", "student_id", "event_id"]].copy()
slot_start = {"Morning": 9.5, "Afternoon": 13.5, "Evening": 16.5}
start_hr = m.loc[att.index, "time_slot"].map(slot_start).values
arrive = start_hr + rng.normal(0, 0.25, len(att))
att["check_in_time"] = [f"{int(h):02d}:{int((h % 1) * 60):02d}" for h in np.clip(arrive, 8, 20)]
att["minutes_stayed"] = np.clip(rng.normal(0.8, 0.22, len(att)) * m.loc[att.index, "duration_hours"].values * 60, 10, None).round().astype(int)
att["full_participation"] = (att["minutes_stayed"] >= 0.8 * m.loc[att.index, "duration_hours"].values * 60).astype(int)

# ------------------------------------------------------------------ feedback
pos = ["Great event, very well organized.", "Learned a lot, loved the speaker.", "Amazing experience, want more events like this.",
       "Super fun and engaging!", "Excellent coordination by the team.", "Very useful and practical session.",
       "Best event of the semester.", "Loved the energy and the prizes.", "Content was top-notch and relevant."]
neu = ["It was okay, nothing special.", "Average event, could be better planned.", "Good idea but the execution was average.",
       "Decent session, some parts were boring.", "Fine overall, venue was a bit crowded.", "Okay experience, timing could improve."]
neg = ["Poorly organized, started very late.", "Too crowded and noisy, could not enjoy it.", "Speaker was not engaging at all.",
       "Waste of time, content was irrelevant.", "Registration process was confusing.", "No food or water arrangements, very disappointing.",
       "Sound system was bad and the event was too long."]

fb_ids = att.index.values
mm = m.loc[fb_ids]
resp_p = np.clip(0.35 + 0.4 * mm["engagement"].values + 0.1 * att["full_participation"].values, 0, 0.95)
responds = rng.random(len(att)) < resp_p
fb = att.loc[responds, ["registration_id", "student_id", "event_id"]].copy()
mf = m.loc[fb.index]
score = (3.7 + 0.55 * mf["event_quality"]
         + 0.15 * mf["has_food"] + 0.12 * mf["has_prizes"]
         - 0.0012 * mf["fee"]
         - 0.07 * np.maximum(mf["duration_hours"] - 4, 0)
         + 0.25 * (mf["guest_speaker_rating"].fillna(3.9) - 3.9)
         - 0.3 * (mf["attendance_probability_true"] if False else 0)
         + rng.normal(0, 0.6, len(fb)))
rating = np.clip(np.round(score), 1, 5).astype(int)
fb["overall_rating"] = rating
fb["organization_rating"] = np.clip(np.round(score + rng.normal(0, 0.5, len(fb))), 1, 5).astype(int)
fb["content_rating"] = np.clip(np.round(score + rng.normal(0, 0.5, len(fb))), 1, 5).astype(int)
fb["venue_rating"] = np.clip(np.round(3.8 + rng.normal(0, 0.8, len(fb)) + 0.3 * (mf["venue"].isin(["Auditorium", "Seminar Hall"]).astype(int))), 1, 5).astype(int)
fb["would_recommend"] = (rating >= 4).astype(int)
fb["sentiment"] = np.where(rating >= 4, "Positive", np.where(rating == 3, "Neutral", "Negative"))
fb["comment"] = [rng.choice(pos) if s == "Positive" else rng.choice(neu) if s == "Neutral" else rng.choice(neg) for s in fb["sentiment"]]
# drop some comments to mimic real missing data
fb.loc[rng.random(len(fb)) < 0.25, "comment"] = np.nan
fb["feedback_id"] = [f"F{str(i).zfill(6)}" for i in range(1, len(fb) + 1)]
fb = fb[["feedback_id", "registration_id", "student_id", "event_id", "overall_rating", "organization_rating",
         "content_rating", "venue_rating", "would_recommend", "sentiment", "comment"]]

# ------------------------------------------------------------ save everything
hidden = ["engagement", "event_quality"]
reg_out = reg.drop(columns=["attendance_probability_true"])
events_out = events.drop(columns=["event_quality"])
student_out = student.drop(columns=["engagement"])
reg_out["registration_date"] = reg_out["registration_date"].dt.date
events_out["event_date"] = events_out["event_date"].dt.date

student_out.to_csv("students.csv", index=False)
events_out.to_csv("events.csv", index=False)
reg_out.to_csv("registrations.csv", index=False)
att.to_csv("attendance.csv", index=False)
fb.to_csv("feedback.csv", index=False)

master = (reg_out.merge(events_out, on="event_id").merge(student_out, on="student_id")
          .merge(fb[["registration_id", "overall_rating", "organization_rating", "content_rating",
                     "venue_rating", "would_recommend", "sentiment", "comment"]], on="registration_id", how="left"))
master.to_csv("master_dataset.csv", index=False)

print(f"students:       {len(student_out):>8,}")
print(f"events:         {len(events_out):>8,}")
print(f"registrations:  {len(reg_out):>8,}")
print(f"attendance:     {len(att):>8,}")
print(f"feedback:       {len(fb):>8,}")
print(f"master rows:    {len(master):>8,}  columns: {master.shape[1]}")
print(f"overall attendance rate: {reg_out['attended'].mean():.1%}")
print(f"mean rating: {fb['overall_rating'].mean():.2f}")

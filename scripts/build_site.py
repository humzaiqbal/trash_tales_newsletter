#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import html
import json
import os
import re
import shutil
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET

try:
    from PIL import Image
except ImportError:  # pragma: no cover - local builds can still fall back to raw DOCX images.
    Image = None


ROOT_DIR = Path(__file__).resolve().parent.parent
SOURCE_DIR = Path(os.environ.get("TRASH_TALES_SOURCE_DIR", str(Path.home() / "Downloads" / "newsletter")))
SITE_DIR = ROOT_DIR / "site"
POSTS_DIR = SITE_DIR / "posts"
ASSETS_DIR = SITE_DIR / "assets"
IMAGES_DIR = ASSETS_DIR / "images"
LOCAL_CHARACTER_LIST_FILE = ROOT_DIR / "Character List.docx"
CHARACTER_LIST_FILE = (
    LOCAL_CHARACTER_LIST_FILE
    if LOCAL_CHARACTER_LIST_FILE.exists()
    else SOURCE_DIR / "Character List.docx"
)
WORKOUT_LOG_CANDIDATES = (
    sorted(ROOT_DIR.glob("*Workout Log.xlsx"))
    or sorted(SOURCE_DIR.glob("*Workout Log.xlsx"))
)
WORKOUT_LOG_FILE = (
    WORKOUT_LOG_CANDIDATES[0]
    if WORKOUT_LOG_CANDIDATES
    else SOURCE_DIR / "Workout Log.xlsx"
)
ASSET_VERSION = "20260710a"
IMAGE_MAX_WIDTH = 1600
IMAGE_WEBP_QUALITY = 84
WORKOUT_PROGRESS_CONFIG = {
    "81": {
        "dates": {dt.date(2026, 6, 23), dt.date(2026, 6, 25)},
        "metric": "volume",
        "description": (
            "Charts below show the logged progress over time for each exercise from the "
            "6/23/2026 and 6/25/2026 workouts. For weighted exercises, the metric is total "
            "volume; for bodyweight, time, or weight-only entries, it uses the clearest "
            "logged quantity available."
        ),
    },
    "82": {
        "dates": {dt.date(2026, 6, 25)},
        "metric": "average_load",
        "description": (
            "Charts below show progress over time for each exercise from the 6/25/2026 "
            "workout. The metric is average load per rep across the logged sets, which "
            "normalizes progress instead of adding all three sets into total volume."
        ),
    },
    "83": {
        "dates": {dt.date(2026, 7, 7), dt.date(2026, 7, 9)},
        "metric": "average_load",
        "description": (
            "Charts below show progress through 7/9/2026 for every exercise from the "
            "7/7/2026 and 7/9/2026 workouts; later dates are excluded. Weighted exercises "
            "use average load per rep across the logged sets. Bodyweight or non-numeric "
            "load entries use average reps or time per set."
        ),
    },
    "84": {
        "dates": {dt.date(2026, 7, 16)},
        "metric": "average_load",
        "description": (
            "Charts below show progress through 7/16/2026 for every exercise from the "
            "7/16/2026 workout; later dates are excluded. The metric is average load per "
            "rep across the logged sets rather than total volume."
        ),
    },
    "85": {
        "dates": {dt.date(2026, 7, 23)},
        "metric": "average_load",
        "description": (
            "Charts below show progress through 7/23/2026 for every exercise from the "
            "7/23/2026 workout. The metric is average load per rep across the logged sets "
            "rather than total volume."
        ),
    },
    "86": {
        "dates": {dt.date(2026, 7, 30)},
        "metric": "average_load",
        "exercise_scope": "through_cutoff",
        "description": (
            "The charts below show progress for every exercise logged through 7/30/2026; "
            "later dates are excluded. Weighted exercises use average load per rep across "
            "the logged sets rather than total volume. Bodyweight exercises use average "
            "reps or time per set."
        ),
        "note": (
            "Because of my ultramarathon attempt, I couldn't do squats or deadlifts "
            "like usual."
        ),
    },
}
WORKOUT_XLSX_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "office": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

QUOTED_NICKNAMES = re.compile(r"[\"“”]([^\"“”]+)[\"“”]")
EPISODE_NUMBER = re.compile(r"episode_(\d+)", re.IGNORECASE)
EPISODE_LABEL = re.compile(r"episode_([0-9]+(?:-[0-9]+)?)", re.IGNORECASE)
PRIVATE_AUTHOR_NAME = (
    re.compile(
        re.escape(WORKOUT_LOG_FILE.stem.removesuffix(" - Workout Log")) + "l?",
        re.IGNORECASE,
    )
    if WORKOUT_LOG_CANDIDATES
    else None
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "a": A_NS, "r": R_NS, "pr": PKG_REL_NS}

SKIP_LINES = {
    "character list",
    "i made a character list detailing every character here. if you are lost as to who someone is look here :d",
    "cool things about people",
    "heres a list of some of the best responses i’ve gotten to the question “what’s something cool or interesting about you that people wouldn’t expect?”",
    "heres a list of some of the best responses i’ve gotten to the question \"what’s something cool or interesting about you that people wouldn’t expect?\"",
}

SPECIAL_TITLES = {
    "62": "Episode 62- Congrats to my CFAs!!",
    "64": "Episode 64 - New website",
}

EXCERPT_SKIP_LINES = SKIP_LINES | {
    "previous episodes",
    "if you want to find previous episodes look here.",
    "if you want to see any of the previous newsletters look here.",
}

EPISODE_TEXT_REPLACEMENTS = {
    "65": {
        "Tightly Knit": "Tight Knit",
        "The Spikeballer": "The Spiker",
        "The spikeballer": "The Spiker",
    },
    "69": {
        'Shoutout to “The Politician” for': "",
        'Shoutout to "The Politician" for': "",
    },
    "70": {
        "newsletter https://www.citysmartnews.com/": "newsletter called city smart https://www.citysmartnews.com/",
        "newsletter (https://www.citysmartnews.com/)": "newsletter called city smart (https://www.citysmartnews.com/)",
    }
}

INLINE_LINKS = [
    (
        "luck surface area",
        "https://www.codusoperandi.com/posts/increasing-your-luck-surface-area",
    ),
    (
        "city smart",
        "https://www.citysmartnews.com/",
    ),
    (
        "meme song",
        "https://www.youtube.com/watch?v=aMhHWWIxK-4&feature=youtu.be",
    ),
    (
        "Assessor Recorder",
        "https://www.sf.gov/departments--assessor-recorder",
    ),
    (
        "a Yelp Review",
        "https://www.yelp.com/biz/kowloon-tong-dessert-cafe-san-francisco?hrid=oyF7m2y0KoziaZhPGojM6A&utm_campaign=www_review_share_popup&utm_medium=copy_link&utm_source=(direct)",
    ),
    (
        "episode 45",
        "https://trashtalesnewsletter.github.io/trash_tales_newsletter/posts/episode-45.html",
    ),
]


FLASHCARDS = [
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What is the formal name of San Francisco?",
        "back": "The City and County of San Francisco.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What are the three levels of government discussed in class?",
        "back": "Federal, state, and local.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "How long are elected officials' terms in San Francisco?",
        "back": "Four years.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "Name the elected officials/entities listed for San Francisco.",
        "back": "Mayor, City Attorney, Public Defender, Sheriff, Board of Supervisors, Assessor-Recorder, City College Board of Trustees, Treasurer, District Attorney, and Board of Education.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What is the simple template for San Francisco's executive branch?",
        "back": "Elected officials appoint commissions, commissions oversee departments, and departments do the work.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "Roughly how many commissions and departments does San Francisco have?",
        "back": "About 51-60 commissions and about 51-60 departments.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What is the approximate total San Francisco government budget for FY25-26?",
        "back": "About $16 billion.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What are the two broad categories of the city budget?",
        "back": "The General Fund and enterprise departments.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What are enterprise departments?",
        "back": "Self-supporting government agencies that generate their own revenue.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "Which groups are social or political entities rather than official government bodies?",
        "back": "Democratic clubs and neighborhood associations.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "Which groups are official advisory entities where citizens can provide input?",
        "back": "Citizens' Advisory Committees and the Civil Grand Jury.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What is the long-term roadmap for the city's physical development, and what law implements it?",
        "back": "The General Plan is the roadmap. The Planning Code is the local statutory law that implements it.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What are examples of mandates and state preemption in housing policy?",
        "back": "California requiring local governments to meet RHNA targets is a mandate. SB 423 overriding local planning is state preemption.",
    },
    {
        "deck": "Class 1: SF Government Basics",
        "front": "What are BART and ABAG?",
        "back": "BART is a regional public transit agency. ABAG is a council of governments that coordinates land use and housing policy across the region.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the Mayor do?",
        "back": "The Mayor is the chief executive and head of the executive branch.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the Board of Supervisors do?",
        "back": "The Board of Supervisors is the legislative body.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the Board of Education oversee?",
        "back": "The San Francisco Unified School District (SFUSD).",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the City College Board of Trustees oversee?",
        "back": "San Francisco's community college.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the Treasurer do?",
        "back": "Serves as banker, tax collector, and investment manager for San Francisco.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the Assessor-Recorder do?",
        "back": "Assesses taxable property and maintains records, including deeds.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the District Attorney do?",
        "back": "Prosecutes crimes on behalf of the people of San Francisco.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the Public Defender do?",
        "back": "Provides legal representation to people who cannot afford it.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the City Attorney do?",
        "back": "Acts as the city's legal department, drafts legislation, and represents the city in court.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What does the Sheriff do?",
        "back": "Manages county jails, oversees court and government-building security, and handles civil enforcement duties such as warrants and eviction orders.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What is constitutional law?",
        "back": "The highest law in a jurisdiction; it outlines the essential form and function of government.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What is statutory law?",
        "back": "Law created by a legislative body, often called statutes at the state and federal levels.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What is administrative law?",
        "back": "Rules or regulations issued by executive agencies that provide detailed implementation of existing statutory law.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What is case law?",
        "back": "Law created by judicial opinion that interprets the meaning and proper application of other types of law.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "How many members are on the Board of Supervisors, and how many members are on standing legislative committees?",
        "back": "The Board of Supervisors has 11 members. Standing legislative committees have 3 members.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "How is the President of the Board of Supervisors chosen?",
        "back": "By a vote of the Supervisors.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "Who assigns membership of legislative committees?",
        "back": "The President of the Board.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What are the three main legislative actions the Board of Supervisors can take?",
        "back": "Ordinances, resolutions, and motions.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What is an ordinance?",
        "back": "A form of statutory law adopted through the legislative process.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What is a resolution?",
        "back": "A declaration of policy, subject to statutory law.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What is a motion?",
        "back": "A procedural act within the Board's exclusive jurisdiction, such as rules of order.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "What may happen when a proposed ordinance is heard by a legislative committee?",
        "back": "The committee may hear expert testimony, receive public comment, vote to make a recommendation, and propose or adopt amendments.",
    },
    {
        "deck": "Class 2: Law and Legislation",
        "front": "Can a legislative committee itself pass an ordinance?",
        "back": "No. It can make recommendations and amendments, but passage occurs through the Board process.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What are the four types of law?",
        "back": "Constitutional law, statutory law, administrative law, and case law.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What is constitutional law?",
        "back": "Law that outlines the essential form and function of government and is the highest law in a jurisdiction.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What is statutory law?",
        "back": "Law created by legislative bodies.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What is administrative law?",
        "back": "Rules and regulations issued by executive agencies.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What is case law?",
        "back": "Law created by judicial opinion that interprets the meaning and application of other types of law.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What are the six major steps of San Francisco's legislative process?",
        "back": "Write legislation, introduce legislation, legislative committee hearing, legislative committee recommendation, full Board vote, and mayoral action.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "Who may cause legislation to be written in the normal legislative process?",
        "back": "Commissions, Supervisors, the Mayor, and departments.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What are the two main ways to put something on the ballot?",
        "back": "Initiative and referral.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What is an initiative?",
        "back": "A ballot measure placed by the people after gathering a sufficient number of signatures.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What is a referral?",
        "back": "A ballot measure placed by the government, with procedures that vary depending on the type of measure.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What types of measures might appear on the ballot?",
        "back": "Ordinances, resolutions, legislative referenda, bond measures, recalls, charter adoption, charter repeal, charter amendments, and charter revisions.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What is the difference between enacted and effective legislation?",
        "back": "Enacted means legislation is formally passed and becomes law on the books. Effective means it becomes enforceable.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "Which city officials may be recalled?",
        "back": "Board of Education members, Supervisors, the Controller, Treasurer, Mayor, District Attorney, City Administrator, Sheriff, and Ethics Commissioners.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "Can a Supervisor send a Charter amendment to the ballot?",
        "back": "Yes, with Board majority approval.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "Can the Mayor send an ordinance to the ballot?",
        "back": "Yes, acting independently.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "How many voter signatures are needed to place a charter amendment on the ballot by petition?",
        "back": "At least 10% of registered voters in San Francisco, which was about 50,000 people as of November 2024.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "Who has the authority to change the San Francisco Charter?",
        "back": "The people.",
    },
    {
        "deck": "Class 3: Ballot Measures and Charter",
        "front": "What are the key dates in the Class 3 political history timeline?",
        "back": "1849: California Constitution adopted. 1850: first State Legislature creates SF County and SF City. 1856: Consolidation Act. 1879: second California Constitution revision. 1898: first SF Charter adopted. 1932: second SF Charter revision. 1996: third SF Charter revision.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What are the General Plan, Housing Element, and Planning Code?",
        "back": "The General Plan is the long-term roadmap for physical development. The Housing Element describes actions to meet the RHNA target. The Planning Code is the local statutory law implementing the roadmap.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "How long is each RHNA cycle?",
        "back": "Eight years.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What is San Francisco's current RHNA target and deadline?",
        "back": "San Francisco must approve 82,000 new housing units by 2031.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "How is the Planning Commission appointed?",
        "back": "Four members are appointed by the Mayor with Board approval, and three are appointed by the Board.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "Which government entity drafts the Housing Element?",
        "back": "The San Francisco Planning Department.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "Which government entity sets the Regional Housing Needs Distribution (RHND)?",
        "back": "The California Department of Housing and Community Development (HCD).",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "Which government entity enforces the Planning Code?",
        "back": "The San Francisco Planning Department.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "Which government entity processes legislative referrals for General Plan and Planning Code consistency?",
        "back": "The San Francisco Planning Department.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What Planning Commission action can allow a project on a parcel not zoned for that use without changing statutory law?",
        "back": "A Conditional Use Authorization.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What is the difference between the RHNA target and SB 423?",
        "back": "The RHNA target is a state-imposed mandate. SB 423 is an example of state preemption overriding certain local land use controls.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "If you get a speeding ticket in San Francisco, which court would you appear before?",
        "back": "San Francisco Superior Court.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What happened in 1849 in the political history timeline?",
        "back": "The California Constitution was adopted.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What happened in 1850 in the political history timeline?",
        "back": "The first State Legislature created San Francisco County and San Francisco City.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What happened in 1856 in the political history timeline?",
        "back": "The Consolidation Act.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What happened in 1879 in the political history timeline?",
        "back": "The second California Constitution was adopted as a revision.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What happened in 1898 in the political history timeline?",
        "back": "The San Francisco Charter was adopted.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "When were the Planning Commission, Planning Code, and Planning Department created?",
        "back": "Planning Commission: 1917. Planning Code: 1921. Planning Department: 1942.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "When was the first General Plan adopted?",
        "back": "1945.",
    },
    {
        "deck": "Class 4: Planning and Housing",
        "front": "What happened in 1960 and 1996 in the political history timeline?",
        "back": "1960: second San Francisco Planning Code. 1996: third San Francisco Charter revision.",
    },
]


FLASHCARDS.extend(
    [
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Which San Francisco government roles are not elected offices?",
            "back": "Clerk of the Board, Superior Court Judge, Controller, City Administrator, Chief of Police, and County Clerk are not elected offices in this civic structure.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Is the Clerk of the Board an elected official in San Francisco?",
            "back": "No.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Is the Controller an elected official in San Francisco?",
            "back": "No.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Is the City Administrator an elected official in San Francisco?",
            "back": "No.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Is the Chief of Police an elected official in San Francisco?",
            "back": "No.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "In the executive branch template, what are commissions described as?",
            "back": "Commissions are the approvers.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "In the executive branch template, what are departments described as?",
            "back": "Departments are the doers.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "In the executive branch template, who appoints commissions?",
            "back": "Elected officials appoint commissions.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "In the executive branch template, who oversees departments?",
            "back": "Commissions oversee departments.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "What is the General Fund?",
            "back": "The primary city budget category, funded mainly through taxes.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Is the General Plan one of the two broad city budget categories?",
            "back": "No. The two broad budget categories are the General Fund and enterprise departments.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Are the Civil Grand Jury and Citizens' Advisory Committees social/political entities or advisory entities?",
            "back": "They are advisory entities.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "Are Democratic clubs and neighborhood associations official advisory entities?",
            "back": "No. They are social/political entities, not official advisory entities.",
        },
        {
            "deck": "Class 1: SF Government Basics",
            "front": "What does RHNA stand for in the housing context of these quizzes?",
            "back": "Regional Housing Needs Allocation.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which elected official is San Francisco's chief executive?",
            "back": "The Mayor.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which elected body is San Francisco's legislative body?",
            "back": "The Board of Supervisors.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which elected entity oversees SFUSD?",
            "back": "The Board of Education.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which elected entity oversees San Francisco's community college?",
            "back": "The City College Board of Trustees.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which official serves as banker, tax collector, and investment manager for San Francisco?",
            "back": "The Treasurer.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which official assesses taxable property and maintains property records including deeds?",
            "back": "The Assessor-Recorder.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which official prosecutes crimes on behalf of the people of San Francisco?",
            "back": "The District Attorney.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which official provides legal representation to people who cannot afford it?",
            "back": "The Public Defender.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which office drafts legislation and represents the city in court?",
            "back": "The City Attorney.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which official manages county jails and handles civil enforcement duties like warrants and eviction orders?",
            "back": "The Sheriff.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Is constitutional law issued by executive agencies?",
            "back": "No. Administrative law is issued by executive agencies.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Is statutory law created by judicial opinion?",
            "back": "No. Case law is created by judicial opinion. Statutory law is created by legislative bodies.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Is administrative law often called a rule or regulation?",
            "back": "Yes. Administrative law is often called a rule or regulation.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which type of law is detailed implementation of existing statutory law?",
            "back": "Administrative law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which type of law interprets the meaning and proper application of other types of law?",
            "back": "Case law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Which type of law is often called statute at the state and federal levels?",
            "back": "Statutory law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "What type of law are commission regulations?",
            "back": "Administrative law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "What type of law is the San Francisco Charter?",
            "back": "Statutory law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "What type of law is the Administrative Code?",
            "back": "Statutory law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "What type of law are Municipal Codes?",
            "back": "Statutory law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "What type of law are Mayoral Executive Directives?",
            "back": "Administrative law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "What type of law is the Police Code?",
            "back": "Statutory law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "What type of law are department regulations?",
            "back": "Administrative law.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Can a proposed ordinance heard by legislative committee receive public comment?",
            "back": "Yes.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Can a proposed ordinance heard by legislative committee receive expert testimony?",
            "back": "Yes.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Can a proposed ordinance heard by legislative committee be amended?",
            "back": "Yes. The committee can propose and adopt amendments.",
        },
        {
            "deck": "Class 2: Law and Legislation",
            "front": "Can a proposed ordinance heard by legislative committee get a recommendation vote?",
            "back": "Yes. The committee can vote to make a recommendation.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "What name completes the four-type law framework: constitutional, statutory, administrative, and ____?",
            "back": "Case law.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "What is Step 1 of San Francisco's legislative process?",
            "back": "Write legislation.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "What is Step 2 of San Francisco's legislative process?",
            "back": "Introduce legislation.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "What is Step 3 of San Francisco's legislative process?",
            "back": "Legislative committee hearing.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "What is Step 4 of San Francisco's legislative process?",
            "back": "Legislative committee recommendation.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "What is Step 5 of San Francisco's legislative process?",
            "back": "Full Board vote.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "What is Step 6 of San Francisco's legislative process?",
            "back": "Mayoral action.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Can citizens directly cause legislation to be written in San Francisco's normal legislative process?",
            "back": "No. The selected answers were commissions, Supervisors, the Mayor, and departments.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Can Citizens' Advisory Committees directly cause legislation to be written in San Francisco's normal legislative process?",
            "back": "No. Legislation can be caused by commissions, Supervisors, the Mayor, and departments.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Are motions ballot measures?",
            "back": "No. Motions are procedural acts, not ballot-measure types.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Are court opinions ballot measures?",
            "back": "No. Court opinions are case law, not ballot measures.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Are mayoral executive directives ballot measures?",
            "back": "No. Mayoral executive directives are administrative actions, not ballot measures.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Are commissioner appointments ballot measures?",
            "back": "No. Commissioner appointments are appointment actions, not ballot measures.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Which San Francisco roles are not recallable officials?",
            "back": "Chief of Police, Director of Planning, Police Commissioner, and County Clerk are not recallable officials in this framework.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Can the people change the San Francisco Charter?",
            "back": "Yes. The people have authority to change the Charter.",
        },
        {
            "deck": "Class 3: Ballot Measures and Charter",
            "front": "Can the Board of Supervisors, Mayor, Rules Committee, or City Attorney independently change the San Francisco Charter?",
            "back": "No. Charter changes ultimately require the people.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What does the Housing Element do?",
            "back": "It describes actions the city will take to meet its RHNA target.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What is the Planning Code?",
            "back": "The local statutory law that implements the General Plan roadmap.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What does the SF Planning Department do in the Housing Element process?",
            "back": "The SF Planning Department drafts the Housing Element.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What does HCD do in the Regional Housing Needs Distribution process?",
            "back": "The California Department of Housing and Community Development sets the RHND.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What does the SF Planning Department do with the Planning Code?",
            "back": "The SF Planning Department enforces the Planning Code.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "Which entity processes legislative referrals for consistency with the General Plan and Planning Code priority policies?",
            "back": "The SF Planning Department.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "Which Planning Commission action can allow a project on a parcel not zoned for that use without changing statutory law?",
            "back": "Conditional Use Authorization.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What does a zoning variance generally address?",
            "back": "A zoning variance addresses relief from specific zoning requirements, but the quiz concept to remember is that Conditional Use Authorization allows certain uses without changing statutory law.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What does a Planning Code Amendment do?",
            "back": "A Planning Code Amendment changes the statutory law.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "Which court handles a local San Francisco speeding ticket?",
            "back": "San Francisco Superior Court.",
        },
        {
            "deck": "Class 4: Planning and Housing",
            "front": "What is the Class 4 timeline order from 1849 through 1996?",
            "back": "1849 CA Constitution; 1850 first State Legislature creates SF County and SF City; 1856 Consolidation Act; 1879 second CA Constitution; 1898 SF Charter; 1917 Planning Commission; 1921 Planning Code; 1932 second SF Charter; 1942 Planning Department; 1945 first General Plan; 1960 second Planning Code; 1996 third SF Charter.",
        },
    ]
)


def canonical_alias(alias: str) -> str:
    return alias.strip().strip('"').strip("“").strip("”").strip()


def normalize_alias_lookup(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def is_skippable_line(line: str) -> bool:
    low = line.strip().lower()
    return low in SKIP_LINES


def is_excerpt_skippable_line(line: str) -> bool:
    low = line.strip().lower()
    return low in EXCERPT_SKIP_LINES


def normalize_episode_text(text: str, episode_label: str) -> str:
    if PRIVATE_AUTHOR_NAME:
        text = PRIVATE_AUTHOR_NAME.sub("Trash Tales", text)
    replacements = EPISODE_TEXT_REPLACEMENTS.get(episode_label, {})
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def render_inline_html(
    line: str,
    variant_lookup: Dict[str, Tuple[str, str]],
    variant_pattern: re.Pattern,
) -> str:
    remaining = line
    rendered: List[str] = []

    while remaining:
        earliest_match = None
        earliest_phrase = None
        earliest_url = None

        for phrase, url in INLINE_LINKS:
            pattern = re.compile(rf"{re.escape(phrase)}\s+\(?{re.escape(url)}\)?")
            match = pattern.search(remaining)
            if not match:
                continue
            if earliest_match is None or match.start() < earliest_match.start():
                earliest_match = match
                earliest_phrase = phrase
                earliest_url = url

        if earliest_match is None:
            rendered.append(inject_character_tooltips(remaining, variant_lookup, variant_pattern))
            break

        before = remaining[:earliest_match.start()]
        rendered.append(inject_character_tooltips(before, variant_lookup, variant_pattern))
        rendered.append(
            f'<a href="{html.escape(earliest_url)}" target="_blank" rel="noreferrer">{html.escape(earliest_phrase)}</a>'
        )
        remaining = remaining[earliest_match.end():]
        if remaining.startswith("  "):
            remaining = " " + remaining.lstrip()

    return "".join(rendered)


def write_optimized_image(raw_image: bytes, output_dir: Path, image_count: int, original_target: str) -> str:
    original_ext = Path(original_target).suffix.lower() or ".png"
    fallback_name = f"img-{image_count:03d}{original_ext}"

    if Image is None:
        (output_dir / fallback_name).write_bytes(raw_image)
        return fallback_name

    try:
        with Image.open(BytesIO(raw_image)) as image:
            image.load()
            if getattr(image, "is_animated", False):
                (output_dir / fallback_name).write_bytes(raw_image)
                return fallback_name

            if image.width > IMAGE_MAX_WIDTH:
                height = round(image.height * (IMAGE_MAX_WIDTH / image.width))
                image = image.resize((IMAGE_MAX_WIDTH, height), Image.Resampling.LANCZOS)

            has_alpha = image.mode in {"RGBA", "LA"} or (
                image.mode == "P" and "transparency" in image.info
            )
            image = image.convert("RGBA" if has_alpha else "RGB")
            optimized_name = f"img-{image_count:03d}.webp"
            image.save(
                output_dir / optimized_name,
                "WEBP",
                quality=IMAGE_WEBP_QUALITY,
                method=6,
            )
            return optimized_name
    except Exception:
        (output_dir / fallback_name).write_bytes(raw_image)
        return fallback_name


def read_docx_blocks(
    docx_path: Path,
    image_output_dir: Path | None = None,
    image_url_prefix: str | None = None,
) -> List[dict]:
    blocks: List[dict] = []
    with zipfile.ZipFile(docx_path) as zf:
        doc_xml = zf.read("word/document.xml")
        rels_map: Dict[str, str] = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            for rel in rels_root.findall("pr:Relationship", NS):
                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target", "")
                if rid and target:
                    rels_map[rid] = target

        root = ET.fromstring(doc_xml)
        image_count = 0

        for para in root.findall(".//w:p", NS):
            text_parts: List[str] = []
            para_images: List[str] = []
            for node in para.iter():
                if node.tag == f"{{{W_NS}}}t":
                    text_parts.append(node.text or "")
                elif node.tag == f"{{{A_NS}}}blip":
                    rid = node.attrib.get(f"{{{R_NS}}}embed")
                    if not rid:
                        continue
                    target = rels_map.get(rid)
                    if not target or not image_output_dir or not image_url_prefix:
                        continue
                    internal_path = "word/" + target.lstrip("/")
                    if internal_path not in zf.namelist():
                        continue
                    image_count += 1
                    filename = write_optimized_image(
                        zf.read(internal_path),
                        image_output_dir,
                        image_count,
                        target,
                    )
                    para_images.append(f"{image_url_prefix}/{filename}")

            text = "".join(text_parts).strip()
            if text:
                blocks.append({"type": "paragraph", "text": text})
            for image_url in para_images:
                blocks.append({"type": "image", "url": image_url})
    return blocks


def read_docx_paragraph_lines(docx_path: Path) -> List[str]:
    blocks = read_docx_blocks(docx_path)
    return [b["text"] for b in blocks if b.get("type") == "paragraph" and b.get("text")]


def parse_character_list(lines: List[str]) -> Dict[str, str]:
    aliases: Dict[str, str] = {}
    for line in lines:
        q_matches = list(QUOTED_NICKNAMES.finditer(line))
        if not q_matches:
            continue
        if not (line.strip().startswith("“") or line.strip().startswith('"')):
            continue

        # Only aliases at the beginning of the line are alias declarations.
        # Later quoted names can appear in the description itself.
        leading_matches = []
        idx = 0
        while idx < len(line) and line[idx].isspace():
            idx += 1
        for m in q_matches:
            if m.start() != idx:
                break
            leading_matches.append(m)
            idx = m.end()
            while idx < len(line) and line[idx].isspace():
                idx += 1

        if not leading_matches:
            continue

        desc = line[leading_matches[-1].end() :].strip(" :-.\t")
        if not desc:
            continue

        for m in leading_matches:
            key = canonical_alias(m.group(1).strip())
            if key and key not in aliases:
                aliases[key] = desc
    return aliases


def expand_alias_candidates(alias: str) -> set[str]:
    candidates = {alias}
    # Allow both "YT" and "Youtube" spellings to map to same character.
    candidates.add(re.sub(r"\byt\b", "youtube", alias, flags=re.IGNORECASE))
    candidates.add(re.sub(r"\byoutube\b", "yt", alias, flags=re.IGNORECASE))
    return {c.strip() for c in candidates if c.strip()}


def build_variant_lookup(character_defs: Dict[str, str]) -> Tuple[Dict[str, Tuple[str, str]], re.Pattern]:
    alias_lookup: Dict[str, Tuple[str, str]] = {}
    for alias, desc in character_defs.items():
        for candidate in expand_alias_candidates(alias):
            norm = normalize_alias_lookup(candidate)
            alias_lookup[norm] = (alias, desc)
            if not candidate.lower().startswith("the "):
                alias_lookup[normalize_alias_lookup(f"the {candidate}")] = (alias, desc)
    pattern = re.compile(r"([“\"])\s*([^\"“”]+?)((?:[’']s)?[,.!?]?)\s*([”\"])", re.IGNORECASE)
    return alias_lookup, pattern


def inject_character_tooltips(
    line: str,
    variant_lookup: Dict[str, Tuple[str, str]],
    variant_pattern: re.Pattern,
) -> str:
    if not line:
        return ""
    if not variant_lookup:
        return html.escape(line)

    out: List[str] = []
    last_idx = 0
    for m in variant_pattern.finditer(line):
        start, end = m.span()
        out.append(html.escape(line[last_idx:start]))
        open_q, alias_txt, suffix, close_q = m.groups()
        alias, desc = variant_lookup.get(normalize_alias_lookup(alias_txt), ("", ""))
        if not alias:
            out.append(html.escape(line[start:end]))
        else:
            visible = f"{open_q}{alias_txt}{suffix}{close_q}"
            out.append(
                f'<span class="character-chip" tabindex="0" '
                f'data-character="{html.escape(alias)}" '
                f'data-description="{html.escape(desc)}">{html.escape(visible)}</span>'
            )
        last_idx = end
    out.append(html.escape(line[last_idx:]))
    return "".join(out)


def block_to_html(block: dict, variant_lookup: Dict[str, Tuple[str, str]], variant_pattern: re.Pattern) -> str:
    if block["type"] == "image":
        return (
            '<figure class="post-image-wrap">'
            f'<img class="post-image" src="{html.escape(block["url"])}" loading="lazy" alt="Newsletter image" />'
            "</figure>"
        )
    line = block["text"].strip()
    if not line or is_skippable_line(line):
        return ""
    lowered = line.lower()
    weekdays = {"monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"}
    if lowered in weekdays:
        return f"<h2>{html.escape(line)}</h2>"
    return f"<p>{render_inline_html(line, variant_lookup, variant_pattern)}</p>"


def parse_episode_number(name: str) -> int:
    m = EPISODE_NUMBER.search(name)
    return int(m.group(1)) if m else -1


def parse_episode_label(name: str) -> str:
    m = EPISODE_LABEL.search(name)
    return m.group(1) if m else "unknown"


def render_episode_sidebar(
    posts: List[dict],
    current_url: str | None = None,
    url_prefix: str = "",
) -> str:
    items = []
    for post in posts:
        href = post["url"]
        if url_prefix and href.startswith("./"):
            href = f"{url_prefix}{href[2:]}"
        elif url_prefix:
            href = f"{url_prefix}{href}"
        is_current = post["url"] == current_url
        active_class = " active" if is_current else ""
        current_attr = ' aria-current="page"' if is_current else ""
        items.append(
            f'<li><a class="episode-link{active_class}" href="{html.escape(href)}"{current_attr}>'
            f"{html.escape(post['title'])}</a></li>"
        )
    items_html = "".join(items)
    return f"""
    <aside class="episode-sidebar" aria-label="Episode navigation">
      <div class="episode-sidebar-inner episode-sidebar-desktop">
        <p class="episode-sidebar-label">Browse</p>
        <h2 class="episode-sidebar-title">Episodes</h2>
        <p class="episode-sidebar-meta">{len(posts)} episodes</p>
        <ol class="episode-list">
          {items_html}
        </ol>
      </div>
      <details class="episode-sidebar-mobile">
        <summary class="episode-sidebar-toggle">
          <span>Episodes</span>
          <span class="episode-sidebar-toggle-meta">{len(posts)} episodes</span>
        </summary>
        <div class="episode-sidebar-mobile-body">
          <ol class="episode-list">
            {items_html}
          </ol>
        </div>
      </details>
    </aside>
"""


def render_site_header(home_href: str, quiz_href: str, active: str = "") -> str:
    quiz_active = " active" if active == "quiz" else ""
    quiz_current = ' aria-current="page"' if active == "quiz" else ""
    return f"""
  <header class="site-header">
    <div class="site-header-inner">
      <a class="brand" href="{home_href}">TRASH TALES</a>
      <nav class="site-nav" aria-label="Main navigation">
        <a class="site-nav-link{quiz_active}" href="{quiz_href}"{quiz_current}>Government class quiz</a>
      </nav>
    </div>
  </header>
"""


def render_post_html(
    title: str,
    article_html: str,
    post_date: str,
    canonical_name: str,
    posts: List[dict],
    current_url: str,
) -> str:
    sidebar_html = render_episode_sidebar(posts, current_url=current_url, url_prefix="../")
    header_html = render_site_header("../index.html", "../government-class-quiz.html")
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{html.escape(title)} | Trash Tales</title>
  <meta name="description" content="Weekly Trash Tales newsletter archive." />
  <link rel="stylesheet" href="../assets/styles.css?v={ASSET_VERSION}" />
</head>
<body>
  {header_html}

  <main class="container page-shell">
    {sidebar_html}
    <article class="post page-main">
      <p class="post-meta">{post_date}</p>
      <h1>{html.escape(title)}</h1>
      <section class="post-content">
        {article_html}
      </section>
      <footer class="post-footer">
        <p>Source file: <code>{html.escape(canonical_name)}</code></p>
      </footer>
    </article>
  </main>

  <div id="tooltip" class="tooltip" role="status" aria-live="polite"></div>
  <script src="../assets/app.js?v={ASSET_VERSION}"></script>
</body>
</html>
"""


def render_post_card(post: dict, summary: str) -> str:
    return f"""
      <article class="post-card">
        <p class="post-meta">{html.escape(post['date'])}</p>
        <h2><a href="{html.escape(post['url'])}">{html.escape(post['title'])}</a></h2>
        <p>{html.escape(summary)}</p>
        <a class="read-more" href="{html.escape(post['url'])}">Read post</a>
      </article>
"""


def render_index_html(posts: List[dict]) -> str:
    cards_html = "".join(render_post_card(p, p["excerpt"]) for p in posts)
    sidebar_html = render_episode_sidebar(posts)
    header_html = render_site_header("./index.html", "./government-class-quiz.html")
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Trash Tales Newsletter</title>
  <meta name="description" content="Weekly Trash Tales newsletter archive." />
  <link rel="stylesheet" href="./assets/styles.css?v={ASSET_VERSION}" />
</head>
<body>
  {header_html}

  <main class="container page-shell">
    {sidebar_html}
    <div class="page-main">
      <section class="hero">
        <h1>Weekly Newsletter Archive</h1>
        <p>
          Notes, stories, and characters from each week. Hover or tap a character name to see who they are.
        </p>
      </section>
      <section class="search-panel">
        <label class="search-label" for="archive-search">Search newsletters</label>
        <input
          id="archive-search"
          class="search-input"
          type="search"
          placeholder="Search keywords across all episodes"
          autocomplete="off"
          spellcheck="false"
        />
        <p class="search-help">Search titles and full article text across the archive.</p>
        <p id="search-status" class="search-status" hidden></p>
      </section>
      <section id="archive-grid" class="archive-grid">
        {cards_html}
      </section>
    </div>
  </main>
  <script src="./assets/app.js?v={ASSET_VERSION}" data-search-index="./assets/search-index.json?v={ASSET_VERSION}"></script>
</body>
</html>
"""


def render_quiz_html() -> str:
    cards_json = json.dumps(FLASHCARDS, ensure_ascii=False)
    deck_options = "".join(
        f'<option value="{html.escape(deck)}">{html.escape(deck)}</option>'
        for deck in sorted({card["deck"] for card in FLASHCARDS})
    )
    header_html = render_site_header("./index.html", "./government-class-quiz.html", active="quiz")
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Government Class Quiz | Trash Tales</title>
  <meta name="description" content="Anki-style flashcards for San Francisco government class notes." />
  <link rel="stylesheet" href="./assets/styles.css?v={ASSET_VERSION}" />
</head>
<body>
  {header_html}

  <main class="container quiz-container">
    <section class="hero quiz-hero">
      <h1>Government Class Quiz</h1>
      <p>Anki-style flashcards pulled from the class quiz reports. Flip each card, mark how it felt, and shuffle when you want a fresh order.</p>
    </section>

    <section class="quiz-controls" aria-label="Quiz controls">
      <label class="search-label" for="deck-filter">Deck</label>
      <select id="deck-filter" class="quiz-select">
        <option value="all">All decks</option>
        {deck_options}
      </select>
      <button id="shuffle-cards" class="quiz-secondary-button" type="button">Shuffle</button>
      <p id="quiz-progress" class="search-help"></p>
    </section>

    <section class="flashcard-panel" aria-live="polite">
      <p id="flashcard-deck" class="post-meta"></p>
      <button id="flashcard" class="flashcard" type="button" aria-label="Flip flashcard">
        <span id="flashcard-side" class="flashcard-side">Question</span>
        <span id="flashcard-text" class="flashcard-text"></span>
        <span class="flashcard-hint">Tap to flip</span>
      </button>
      <div class="quiz-actions">
        <button id="again-card" class="quiz-button" type="button">Again</button>
        <button id="good-card" class="quiz-button" type="button">Good</button>
        <button id="easy-card" class="quiz-button" type="button">Easy</button>
      </div>
    </section>

    <section class="quiz-list-wrap">
      <h2>All Cards</h2>
      <div id="quiz-card-list" class="quiz-card-list"></div>
    </section>
  </main>

  <script>
    window.GOVERNMENT_FLASHCARDS = {cards_json};
  </script>
  <script src="./assets/app.js?v={ASSET_VERSION}"></script>
</body>
</html>
"""


def write_assets() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    styles = """* {
  box-sizing: border-box;
}

:root {
  --bg: #fcfcfa;
  --text: #1f2328;
  --muted: #59636e;
  --line: #d7dde4;
  --accent: #0b57d0;
  --chip: #eef3ff;
}

html, body {
  margin: 0;
  padding: 0;
  background: var(--bg);
  color: var(--text);
  font-family: "Inter", "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  line-height: 1.7;
}

.site-header {
  border-bottom: 1px solid var(--line);
  position: sticky;
  top: 0;
  background: color-mix(in oklab, var(--bg) 92%, white 8%);
  backdrop-filter: blur(4px);
  z-index: 20;
}

.site-header-inner {
  max-width: 1180px;
  margin: 0 auto;
  padding: 14px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 20px;
}

.brand {
  color: var(--text);
  text-decoration: none;
  font-weight: 700;
  letter-spacing: 0.08em;
  font-size: 0.9rem;
}

.site-nav {
  display: flex;
  align-items: center;
  gap: 14px;
}

.site-nav-link {
  color: var(--muted);
  text-decoration: none;
  font-size: 0.95rem;
  font-weight: 600;
}

.site-nav-link:hover,
.site-nav-link:focus-visible,
.site-nav-link.active {
  color: var(--text);
}

.container {
  max-width: 1180px;
  margin: 0 auto;
  padding: 30px 20px 64px;
}

.page-shell {
  display: grid;
  grid-template-columns: minmax(170px, 220px) minmax(0, 1fr);
  gap: 24px;
  align-items: start;
}

.page-main {
  min-width: 0;
}

.episode-sidebar {
  align-self: start;
}

.episode-sidebar-inner {
  position: sticky;
  top: 84px;
  max-height: calc(100vh - 104px);
  overflow: auto;
  border: 1px solid var(--line);
  border-radius: 12px;
  padding: 16px;
  background: color-mix(in oklab, white 72%, var(--bg) 28%);
}

.episode-sidebar-mobile {
  display: none;
  border: 1px solid var(--line);
  border-radius: 12px;
  background: color-mix(in oklab, white 72%, var(--bg) 28%);
}

.episode-sidebar-toggle {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 14px 16px;
  cursor: pointer;
  font-weight: 600;
  list-style: none;
}

.episode-sidebar-toggle::-webkit-details-marker {
  display: none;
}

.episode-sidebar-toggle::after {
  content: "+";
  color: var(--muted);
  font-size: 1.1rem;
  line-height: 1;
}

.episode-sidebar-mobile[open] .episode-sidebar-toggle::after {
  content: "-";
}

.episode-sidebar-toggle-meta {
  margin-left: auto;
  padding-right: 8px;
  color: var(--muted);
  font-weight: 500;
  font-size: 0.92rem;
}

.episode-sidebar-mobile-body {
  padding: 0 12px 12px;
}

.episode-sidebar-label {
  margin: 0 0 2px;
  color: var(--muted);
  font-size: 0.78rem;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.episode-sidebar-title {
  margin: 0;
  font-size: 1.05rem;
}

.episode-sidebar-meta {
  margin: 2px 0 14px;
  color: var(--muted);
  font-size: 0.92rem;
}

.episode-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  gap: 6px;
}

.episode-link {
  display: block;
  padding: 8px 10px;
  border-radius: 8px;
  color: var(--muted);
  text-decoration: none;
  overflow-wrap: anywhere;
  transition: background 120ms ease, color 120ms ease;
}

.episode-link:hover,
.episode-link:focus-visible {
  background: #fff;
  color: var(--text);
  outline: none;
}

.episode-link.active {
  background: #fff;
  color: var(--text);
  font-weight: 600;
  box-shadow: inset 0 0 0 1px var(--line);
}

.hero h1,
.post h1 {
  font-family: "Georgia", "Times New Roman", serif;
  font-weight: 700;
  line-height: 1.2;
  letter-spacing: -0.01em;
}

.hero h1 {
  margin: 0 0 10px;
  font-size: clamp(1.9rem, 3.2vw, 2.8rem);
}

.hero p {
  margin: 0 0 24px;
  color: var(--muted);
  max-width: 680px;
}

.search-panel {
  margin: 0 0 24px;
}

.search-label {
  display: block;
  margin: 0 0 8px;
  font-weight: 600;
  color: var(--text);
}

.search-input {
  width: 100%;
  padding: 13px 14px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: #fff;
  color: var(--text);
  font: inherit;
}

.search-input:focus {
  outline: 2px solid color-mix(in oklab, var(--accent) 30%, white 70%);
  border-color: var(--accent);
}

.search-help,
.search-status {
  margin: 8px 0 0;
  color: var(--muted);
  font-size: 0.95rem;
}

.quiz-container {
  max-width: 860px;
}

.quiz-hero p {
  max-width: 720px;
}

.quiz-controls {
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 10px 12px;
  align-items: end;
  margin: 0 0 18px;
}

.quiz-controls .search-label,
.quiz-controls .search-help {
  grid-column: 1 / -1;
}

.quiz-select {
  width: 100%;
  padding: 12px 14px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: #fff;
  color: var(--text);
  font: inherit;
}

.flashcard-panel {
  border: 1px solid var(--line);
  border-radius: 14px;
  padding: 18px;
  background: #fff;
}

.flashcard {
  width: 100%;
  min-height: 260px;
  border: 1px solid var(--line);
  border-radius: 14px;
  padding: 24px;
  background: var(--bg);
  color: var(--text);
  text-align: left;
  cursor: pointer;
  display: grid;
  align-content: center;
  gap: 14px;
}

.flashcard-side {
  color: var(--muted);
  font-size: 0.82rem;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.flashcard-text {
  font-family: "Georgia", "Times New Roman", serif;
  font-size: clamp(1.25rem, 2.5vw, 1.75rem);
  line-height: 1.35;
}

.flashcard-hint {
  color: var(--muted);
  font-size: 0.9rem;
}

.quiz-actions {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin-top: 14px;
}

.quiz-button,
.quiz-secondary-button {
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 11px 14px;
  background: #fff;
  color: var(--text);
  font: inherit;
  font-weight: 600;
  cursor: pointer;
}

.quiz-button:hover,
.quiz-button:focus-visible,
.quiz-secondary-button:hover,
.quiz-secondary-button:focus-visible {
  border-color: var(--accent);
  outline: none;
}

.quiz-list-wrap {
  margin-top: 28px;
}

.quiz-list-wrap h2 {
  margin: 0 0 12px;
  font-family: "Georgia", "Times New Roman", serif;
}

.quiz-card-list {
  display: grid;
  gap: 10px;
}

.quiz-card-summary {
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 14px;
  background: #fff;
}

.quiz-card-summary h3 {
  margin: 0 0 8px;
  font-size: 1rem;
}

.quiz-card-summary p {
  margin: 0;
  color: var(--muted);
}

.archive-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 14px;
}

.post-card {
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 16px;
  background: #fff;
}

.post-card h2 {
  margin: 0 0 8px;
  font-family: "Georgia", "Times New Roman", serif;
  font-size: clamp(1.25rem, 2.2vw, 1.6rem);
}

.post-card h2 a {
  color: var(--text);
  text-decoration: none;
}

.post-card p {
  margin: 0 0 10px;
  color: var(--muted);
}

.post-meta {
  color: var(--muted);
  font-size: 0.88rem;
  letter-spacing: 0.01em;
  margin: 0 0 10px;
}

.read-more {
  color: var(--accent);
  text-decoration: none;
  font-weight: 600;
}

.post h1 {
  margin: 0 0 6px;
  font-size: clamp(2rem, 3.6vw, 3rem);
}

.post-content {
  font-family: "Georgia", "Times New Roman", serif;
  font-size: clamp(1.06rem, 1.4vw, 1.18rem);
}

.post-content p,
.post-content h2,
.post-content h3 {
  max-width: 74ch;
}

.post-content p {
  margin: 0 0 1.15em;
}

.post-content h2 {
  margin: 1.8em 0 0.65em;
  font-size: 1.25em;
}

.post-image-wrap {
  margin: 1.2em 0 1.5em;
}

.post-image {
  display: block;
  width: 100%;
  max-width: 740px;
  height: auto;
  border-radius: 10px;
  border: 1px solid var(--line);
  background: #fff;
}

.workout-progress {
  margin-top: 2.2em;
}

.workout-chart-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(min(100%, 300px), 1fr));
  gap: 16px;
  margin-top: 16px;
}

.workout-chart-card {
  margin: 0;
  border: 1px solid var(--line);
  border-radius: 12px;
  padding: 12px;
  background: #fff;
}

.workout-chart {
  display: block;
  width: 100%;
  height: auto;
}

.workout-chart-card figcaption {
  margin-top: 8px;
  color: var(--muted);
  font-family: "Inter", "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 0.9rem;
}

.post-footer {
  margin-top: 28px;
  padding-top: 10px;
  border-top: 1px solid var(--line);
  color: var(--muted);
  font-size: 0.9rem;
}

.character-chip {
  background: var(--chip);
  border-bottom: 1px dashed var(--accent);
  border-radius: 4px;
  padding: 0 2px;
  cursor: help;
  transition: background 120ms ease;
}

.character-chip:hover,
.character-chip:focus-visible,
.character-chip.active {
  background: #dae6ff;
  outline: none;
}

.tooltip {
  position: fixed;
  z-index: 30;
  max-width: min(320px, calc(100vw - 24px));
  background: #121821;
  color: #f5f8ff;
  border-radius: 8px;
  padding: 10px 12px;
  box-shadow: 0 12px 30px rgba(0, 0, 0, 0.3);
  font-size: 0.9rem;
  line-height: 1.45;
  pointer-events: none;
  opacity: 0;
  transform: translateY(4px);
  transition: opacity 120ms ease, transform 120ms ease;
}

.tooltip.show {
  opacity: 1;
  transform: translateY(0);
}

@media (max-width: 760px) {
  .page-shell {
    grid-template-columns: minmax(150px, 190px) minmax(0, 1fr);
    gap: 16px;
  }

  .episode-sidebar-inner {
    padding: 12px;
  }

  .episode-link {
    padding: 7px 8px;
    font-size: 0.95rem;
  }
}

@media (max-width: 560px) {
  .page-shell {
    grid-template-columns: 1fr;
    gap: 20px;
  }

  .episode-sidebar-desktop {
    display: none;
  }

  .episode-sidebar-mobile {
    display: block;
  }

  .episode-sidebar-mobile .episode-list {
    max-height: min(55vh, 360px);
    overflow: auto;
    padding-right: 2px;
  }

  .episode-sidebar-mobile .episode-link {
    background: #fff;
    box-shadow: inset 0 0 0 1px var(--line);
  }
}

@media (max-width: 640px) {
  .site-header-inner,
  .container {
    padding-left: 14px;
    padding-right: 14px;
  }

  .post-card {
    border-radius: 8px;
    padding: 14px;
  }

  .tooltip {
    font-size: 0.86rem;
  }

  .site-header-inner {
    align-items: flex-start;
    flex-direction: column;
    gap: 8px;
  }

  .quiz-controls,
  .quiz-actions {
    grid-template-columns: 1fr;
  }
}
"""
    (ASSETS_DIR / "styles.css").write_text(styles, encoding="utf-8")

    app_js = """(() => {
  const tooltip = document.getElementById("tooltip");
  if (!tooltip) return;

  let active = null;

  function positionTip(clientX, clientY) {
    const pad = 12;
    const rect = tooltip.getBoundingClientRect();
    let left = clientX + 12;
    let top = clientY + 14;

    if (left + rect.width > window.innerWidth - pad) {
      left = window.innerWidth - rect.width - pad;
    }
    if (top + rect.height > window.innerHeight - pad) {
      top = clientY - rect.height - 14;
    }
    if (left < pad) left = pad;
    if (top < pad) top = pad;

    tooltip.style.left = left + "px";
    tooltip.style.top = top + "px";
  }

  function showTip(el, x, y) {
    const character = el.dataset.character || "Character";
    const description = el.dataset.description || "";
    tooltip.replaceChildren();
    const strong = document.createElement("strong");
    strong.textContent = character;
    const br = document.createElement("br");
    const text = document.createTextNode(description);
    tooltip.append(strong, br, text);
    tooltip.classList.add("show");
    el.classList.add("active");
    active = el;
    positionTip(x, y);
  }

  function hideTip() {
    tooltip.classList.remove("show");
    if (active) active.classList.remove("active");
    active = null;
  }

  function wire(el) {
    el.addEventListener("mouseenter", (e) => showTip(el, e.clientX, e.clientY));
    el.addEventListener("mousemove", (e) => positionTip(e.clientX, e.clientY));
    el.addEventListener("mouseleave", hideTip);

    el.addEventListener("focus", () => {
      const r = el.getBoundingClientRect();
      showTip(el, r.left + r.width / 2, r.top + r.height / 2);
    });
    el.addEventListener("blur", hideTip);

    el.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const r = el.getBoundingClientRect();
      if (active === el) {
        hideTip();
      } else {
        showTip(el, r.left + r.width / 2, r.top + r.height / 2);
      }
    });
  }

  document.querySelectorAll(".character-chip").forEach(wire);
  document.addEventListener("click", (e) => {
    if (!(e.target instanceof Element)) return;
    if (!e.target.closest(".character-chip")) hideTip();
  });
  window.addEventListener("scroll", () => {
    if (active) {
      const r = active.getBoundingClientRect();
      positionTip(r.left + r.width / 2, r.top + r.height / 2);
    }
  }, { passive: true });
})();

(() => {
  const input = document.getElementById("archive-search");
  const grid = document.getElementById("archive-grid");
  const status = document.getElementById("search-status");
  if (!input || !grid || !status) return;

  const scriptEl = document.querySelector("script[data-search-index]");
  const searchIndexUrl = scriptEl?.dataset.searchIndex;
  if (!searchIndexUrl) return;

  const initialHtml = grid.innerHTML;
  let posts = [];
  let loaded = false;

  function escapeHtml(value) {
    return value
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function normalize(value) {
    return value.toLowerCase().replace(/\s+/g, " ").trim();
  }

  async function loadSearchIndex() {
    if (loaded) return posts;
    const response = await fetch(searchIndexUrl);
    if (!response.ok) {
      throw new Error("Could not load search index");
    }
    posts = await response.json();
    loaded = true;
    return posts;
  }

  function buildSnippet(post, queryTerms) {
    const source = post.content || post.excerpt || "";
    const lowerSource = source.toLowerCase();
    let matchIndex = -1;

    for (const term of queryTerms) {
      const idx = lowerSource.indexOf(term);
      if (idx !== -1 && (matchIndex === -1 || idx < matchIndex)) {
        matchIndex = idx;
      }
    }

    if (matchIndex === -1) {
      return post.excerpt || source.slice(0, 240);
    }

    const start = Math.max(0, matchIndex - 70);
    const end = Math.min(source.length, matchIndex + 170);
    let snippet = source.slice(start, end).trim();
    if (start > 0) snippet = "..." + snippet;
    if (end < source.length) snippet += "...";
    return snippet;
  }

  function renderCards(results, queryTerms) {
    grid.innerHTML = results.map((post) => `
      <article class="post-card">
        <p class="post-meta">${escapeHtml(post.date)}</p>
        <h2><a href="${escapeHtml(post.url)}">${escapeHtml(post.title)}</a></h2>
        <p>${escapeHtml(buildSnippet(post, queryTerms))}</p>
        <a class="read-more" href="${escapeHtml(post.url)}">Read post</a>
      </article>
    `).join("");
  }

  function setStatus(message, hidden = false) {
    status.textContent = message;
    status.hidden = hidden;
  }

  async function runSearch() {
    const query = normalize(input.value);
    if (!query) {
      grid.innerHTML = initialHtml;
      setStatus("", true);
      return;
    }

    try {
      const data = await loadSearchIndex();
      const terms = query.split(" ").filter(Boolean);
      const results = data.filter((post) => {
        const haystack = normalize(`${post.title} ${post.content}`);
        return terms.every((term) => haystack.includes(term));
      });

      renderCards(results, terms);
      const label = results.length === 1 ? "result" : "results";
      setStatus(`${results.length} ${label} for "${input.value.trim()}"`);
    } catch (error) {
      setStatus("Search is temporarily unavailable.");
    }
  }

  input.addEventListener("input", runSearch);
})();

(() => {
  const allCards = window.GOVERNMENT_FLASHCARDS || [];
  const cardButton = document.getElementById("flashcard");
  const cardText = document.getElementById("flashcard-text");
  const cardSide = document.getElementById("flashcard-side");
  const cardDeck = document.getElementById("flashcard-deck");
  const deckFilter = document.getElementById("deck-filter");
  const progress = document.getElementById("quiz-progress");
  const list = document.getElementById("quiz-card-list");
  const shuffleButton = document.getElementById("shuffle-cards");
  const reviewButtons = [
    document.getElementById("again-card"),
    document.getElementById("good-card"),
    document.getElementById("easy-card"),
  ];

  if (!allCards.length || !cardButton || !cardText || !cardSide || !cardDeck || !deckFilter || !progress || !list) return;

  let cards = [...allCards];
  let index = 0;
  let showingBack = false;

  function escapeHtml(value) {
    return value
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function filteredCards() {
    const selectedDeck = deckFilter.value;
    if (selectedDeck === "all") return [...allCards];
    return allCards.filter((card) => card.deck === selectedDeck);
  }

  function renderList() {
    list.innerHTML = cards.map((card) => `
      <article class="quiz-card-summary">
        <p class="post-meta">${escapeHtml(card.deck)}</p>
        <h3>${escapeHtml(card.front)}</h3>
        <p>${escapeHtml(card.back)}</p>
      </article>
    `).join("");
  }

  function renderCard() {
    if (!cards.length) {
      cardDeck.textContent = "";
      cardSide.textContent = "No cards";
      cardText.textContent = "No cards match this deck.";
      progress.textContent = "";
      list.innerHTML = "";
      return;
    }

    const card = cards[index];
    cardDeck.textContent = card.deck;
    cardSide.textContent = showingBack ? "Answer" : "Question";
    cardText.textContent = showingBack ? card.back : card.front;
    progress.textContent = `Card ${index + 1} of ${cards.length}`;
    renderList();
  }

  function nextCard() {
    if (!cards.length) return;
    index = (index + 1) % cards.length;
    showingBack = false;
    renderCard();
  }

  function shuffleCards() {
    for (let i = cards.length - 1; i > 0; i -= 1) {
      const j = Math.floor(Math.random() * (i + 1));
      [cards[i], cards[j]] = [cards[j], cards[i]];
    }
    index = 0;
    showingBack = false;
    renderCard();
  }

  cardButton.addEventListener("click", () => {
    showingBack = !showingBack;
    renderCard();
  });

  deckFilter.addEventListener("change", () => {
    cards = filteredCards();
    index = 0;
    showingBack = false;
    renderCard();
  });

  shuffleButton?.addEventListener("click", shuffleCards);
  reviewButtons.forEach((button) => button?.addEventListener("click", nextCard));

  renderCard();
})();
"""
    (ASSETS_DIR / "app.js").write_text(app_js, encoding="utf-8")


def excerpt_from_lines(lines: List[str], max_len: int = 260) -> str:
    cleaned_lines = [line for line in lines if line and not is_excerpt_skippable_line(line)]
    joined = " ".join(cleaned_lines[:12]).strip()
    joined = re.sub(r"\s+", " ", joined)
    if len(joined) <= max_len:
        return joined
    return joined[: max_len - 1].rstrip() + "…"


def xlsx_column_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return -1
    index = 0
    for char in match.group(1):
        index = index * 26 + ord(char) - ord("A") + 1
    return index


def excel_serial_to_date(value: object) -> dt.date | None:
    try:
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(value))).date()
    except (TypeError, ValueError):
        return None


def read_xlsx_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("main:si", WORKOUT_XLSX_NS):
        strings.append("".join(t.text or "" for t in item.findall(".//main:t", WORKOUT_XLSX_NS)))
    return strings


def read_xlsx_cell_value(cell: ET.Element, shared_strings: List[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("main:v", WORKOUT_XLSX_NS)
    if cell_type == "s" and value is not None:
        return shared_strings[int(value.text or "0")]
    if cell_type == "inlineStr":
        inline = cell.find("main:is", WORKOUT_XLSX_NS)
        if inline is not None:
            return "".join(t.text or "" for t in inline.findall(".//main:t", WORKOUT_XLSX_NS))
    return value.text if value is not None and value.text is not None else ""


def parse_numeric(value: object) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text in {"-", "BW"}:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


def workout_metric(
    sets: List[Tuple[str, str]],
    metric_mode: str = "volume",
) -> Tuple[float | None, str]:
    volume = 0.0
    has_volume = False
    reps_total = 0.0
    weights = []
    weighted_reps = 0.0
    for reps, weight in sets:
        reps_num = parse_numeric(reps)
        weight_num = parse_numeric(weight)
        if reps_num is not None:
            reps_total += reps_num
        if weight_num is not None:
            weights.append(weight_num)
        if reps_num is not None and weight_num is not None:
            volume += reps_num * weight_num
            weighted_reps += reps_num
            has_volume = True

    if metric_mode == "average_load":
        if has_volume and weighted_reps:
            return volume / weighted_reps, "average load per rep"
        if weights:
            return sum(weights) / len(weights), "average logged load"
        return None, "average load per rep"

    if metric_mode == "average_reps":
        numeric_sets = [
            parse_numeric(reps)
            for reps, _ in sets
            if parse_numeric(reps) is not None
        ]
        if numeric_sets:
            return sum(numeric_sets) / len(numeric_sets), "average reps/time per set"
        return None, "average reps/time per set"

    if has_volume:
        return volume, "total volume"
    if reps_total:
        return reps_total, "total reps/time"
    if weights:
        return max(weights), "max logged weight"
    return None, "logged work"


def read_workout_rows(workout_path: Path) -> List[dict]:
    if not workout_path.exists():
        return []

    with zipfile.ZipFile(workout_path) as zf:
        shared_strings = read_xlsx_shared_strings(zf)
        root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        current_date = None
        rows = []
        for row in root.findall(".//main:sheetData/main:row", WORKOUT_XLSX_NS):
            values: Dict[int, str] = {}
            for cell in row.findall("main:c", WORKOUT_XLSX_NS):
                col = xlsx_column_index(cell.attrib.get("r", ""))
                if col > 0:
                    values[col] = read_xlsx_cell_value(cell, shared_strings)

            if values.get(1):
                current_date = excel_serial_to_date(values[1])
            exercise = values.get(2, "").strip()
            if not current_date or not exercise or exercise.lower() == "exercise":
                continue

            sets = []
            for reps_col, weight_col in [(3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]:
                reps = values.get(reps_col, "").strip()
                weight = values.get(weight_col, "").strip()
                if reps or weight:
                    sets.append((reps, weight))
            metric, metric_label = workout_metric(sets)
            if metric is None:
                continue
            rows.append(
                {
                    "date": current_date,
                    "exercise": exercise,
                    "metric": metric,
                    "metric_label": metric_label,
                    "sets": sets,
                }
            )
        return rows


def slugify_chart_name(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
    return slug or "exercise"


def render_workout_svg(exercise: str, metric_label: str, points: List[Tuple[dt.date, float]]) -> str:
    width = 760
    height = 360
    left = 72
    right = width - 28
    top = 38
    bottom = height - 58
    values = [value for _, value in points]
    min_value = min(values)
    max_value = max(values)
    if min_value == max_value:
        pad = max(1.0, max_value * 0.1)
        min_value -= pad
        max_value += pad
    else:
        pad = (max_value - min_value) * 0.12
        min_value -= pad
        max_value += pad

    def x_for(index: int) -> float:
        if len(points) == 1:
            return (left + right) / 2
        return left + ((right - left) * index / (len(points) - 1))

    def y_for(value: float) -> float:
        return bottom - ((value - min_value) * (bottom - top) / (max_value - min_value))

    coordinates = [(x_for(i), y_for(value), date, value) for i, (date, value) in enumerate(points)]
    path = " ".join(f"{x:.1f},{y:.1f}" for x, y, _, _ in coordinates)
    circles = "\n".join(
        f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4.5"><title>{date.strftime("%b %-d")}: {value:g}</title></circle>'
        for x, y, date, value in coordinates
    )
    first_date = points[0][0].strftime("%b %-d")
    last_date = points[-1][0].strftime("%b %-d")
    y_ticks = [min_value, (min_value + max_value) / 2, max_value]
    y_labels = "\n".join(
        f'<text x="{left - 10}" y="{y_for(value) + 4:.1f}" text-anchor="end">{value:g}</text>'
        for value in y_ticks
    )

    return f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width} {height}" role="img" aria-labelledby="title desc">
  <title>{html.escape(exercise)} progress</title>
  <desc>{html.escape(metric_label)} over time from {html.escape(first_date)} to {html.escape(last_date)}.</desc>
  <rect width="100%" height="100%" fill="#ffffff"/>
  <text x="{left}" y="24" font-family="Inter, Arial, sans-serif" font-size="18" font-weight="700" fill="#1f2328">{html.escape(exercise)}</text>
  <text x="{left}" y="44" font-family="Inter, Arial, sans-serif" font-size="12" fill="#59636e">{html.escape(metric_label)}</text>
  <line x1="{left}" y1="{bottom}" x2="{right}" y2="{bottom}" stroke="#d7dde4"/>
  <line x1="{left}" y1="{top}" x2="{left}" y2="{bottom}" stroke="#d7dde4"/>
  <line x1="{left}" y1="{y_for(y_ticks[1]):.1f}" x2="{right}" y2="{y_for(y_ticks[1]):.1f}" stroke="#eef1f5"/>
  {y_labels}
  <polyline points="{path}" fill="none" stroke="#0b57d0" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
  <g fill="#0b57d0">{circles}</g>
  <text x="{left}" y="{height - 24}" font-family="Inter, Arial, sans-serif" font-size="12" fill="#59636e" text-anchor="start">{html.escape(first_date)}</text>
  <text x="{right}" y="{height - 24}" font-family="Inter, Arial, sans-serif" font-size="12" fill="#59636e" text-anchor="end">{html.escape(last_date)}</text>
</svg>
"""


def render_workout_progress_section(
    episode_label: str,
    image_output_dir: Path,
    image_url_prefix: str,
) -> str:
    config = WORKOUT_PROGRESS_CONFIG.get(episode_label)
    if not config:
        return ""

    rows = read_workout_rows(WORKOUT_LOG_FILE)
    if not rows:
        return ""

    progress_cutoff = max(config["dates"])
    include_all_through_cutoff = config.get("exercise_scope") == "through_cutoff"
    target_exercises = []
    seen = set()
    for row in rows:
        exercise = row["exercise"]
        is_target_row = (
            row["date"] <= progress_cutoff
            if include_all_through_cutoff
            else row["date"] in config["dates"]
        )
        if is_target_row and exercise not in seen:
            target_exercises.append(exercise)
            seen.add(exercise)

    if not target_exercises:
        return ""

    chart_items = []
    for exercise in target_exercises:
        by_date: Dict[dt.date, float] = {}
        metric_label = "logged work"
        metric_mode = config["metric"]
        if metric_mode == "average_load":
            target_rows = [
                row
                for row in rows
                if row["exercise"] == exercise
                and (
                    row["date"] <= progress_cutoff
                    if include_all_through_cutoff
                    else row["date"] in config["dates"]
                )
            ]
            if not any(
                workout_metric(row["sets"], "average_load")[0] is not None
                for row in target_rows
            ):
                metric_mode = "average_reps"
        for row in rows:
            if row["exercise"] != exercise or row["date"] > progress_cutoff:
                continue
            metric, row_metric_label = workout_metric(row["sets"], metric_mode)
            if metric is None:
                continue
            metric_label = row_metric_label
            by_date[row["date"]] = max(by_date.get(row["date"], float("-inf")), metric)
        points = sorted(by_date.items())
        if not points:
            continue
        filename = f"workout-{slugify_chart_name(exercise)}.svg"
        (image_output_dir / filename).write_text(
            render_workout_svg(exercise, metric_label, points),
            encoding="utf-8",
        )
        chart_items.append(
            f"""
        <figure class="workout-chart-card">
          <img class="workout-chart" src="{html.escape(image_url_prefix)}/{html.escape(filename)}" loading="lazy" alt="{html.escape(exercise)} progress chart" />
          <figcaption>{html.escape(exercise)} — {html.escape(metric_label)}</figcaption>
        </figure>"""
        )

    if not chart_items:
        return ""

    note_html = (
        f'\n        <p class="workout-note">* {html.escape(config["note"])}</p>'
        if config.get("note")
        else ""
    )
    return f"""
      <section class="workout-progress">
        <h2>Personal training progress for the week</h2>
        <p>{html.escape(config["description"])}</p>{note_html}
        <div class="workout-chart-grid">
          {"".join(chart_items)}
        </div>
      </section>
"""


def build() -> None:
    SITE_DIR.mkdir(parents=True, exist_ok=True)
    POSTS_DIR.mkdir(parents=True, exist_ok=True)
    write_assets()
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)

    character_lines = read_docx_paragraph_lines(CHARACTER_LIST_FILE)
    character_defs = parse_character_list(character_lines)
    variant_lookup, variant_pattern = build_variant_lookup(character_defs)

    episode_by_label = {}
    for source_dir in (SOURCE_DIR, ROOT_DIR):
        for path in sorted(source_dir.glob("episode_*.docx")):
            episode_by_label[parse_episode_label(path.name)] = path
    episode_files = list(episode_by_label.values())
    episode_files = sorted(episode_files, key=lambda p: parse_episode_number(p.name), reverse=True)

    index_posts = []
    for docx_path in episode_files:
        episode_label = parse_episode_label(docx_path.name)
        title = SPECIAL_TITLES.get(episode_label, f"Episode {episode_label}")
        post_slug = f"episode-{episode_label}" if episode_label != "unknown" else docx_path.stem
        post_url = f"./posts/{post_slug}.html"
        post_image_dir = IMAGES_DIR / post_slug
        if post_image_dir.exists():
            shutil.rmtree(post_image_dir)
        post_image_dir.mkdir(parents=True, exist_ok=True)

        blocks = read_docx_blocks(
            docx_path,
            image_output_dir=post_image_dir,
            image_url_prefix=f"../assets/images/{post_slug}",
        )
        for block in blocks:
            if block.get("type") == "paragraph":
                block["text"] = normalize_episode_text(block["text"], episode_label)
        lines = [b["text"] for b in blocks if b.get("type") == "paragraph"]
        article_blocks = [block_to_html(b, variant_lookup, variant_pattern) for b in blocks]
        article_html = "\n        ".join([b for b in article_blocks if b])
        if episode_label in WORKOUT_PROGRESS_CONFIG:
            workout_html = render_workout_progress_section(
                episode_label,
                post_image_dir,
                f"../assets/images/{post_slug}",
            )
            if workout_html:
                article_html = f"{article_html}\n        {workout_html}"

        index_posts.append(
            {
                "title": title,
                "date": dt.datetime.fromtimestamp(docx_path.stat().st_mtime).strftime("%b %d, %Y"),
                "url": post_url,
                "excerpt": excerpt_from_lines(lines),
                "content": " ".join(
                    line for line in lines if line and not is_excerpt_skippable_line(line)
                ),
                "article_html": article_html,
                "canonical_name": docx_path.name,
            }
        )

    for post in index_posts:
        post_slug = Path(post["url"]).stem
        page = render_post_html(
            post["title"],
            post["article_html"],
            post["date"],
            post["canonical_name"],
            index_posts,
            post["url"],
        )
        (POSTS_DIR / f"{post_slug}.html").write_text(page, encoding="utf-8")

    (ASSETS_DIR / "search-index.json").write_text(
        json.dumps(
            [
                {
                    "title": post["title"],
                    "date": post["date"],
                    "url": post["url"],
                    "excerpt": post["excerpt"],
                    "content": post["content"],
                }
                for post in index_posts
            ],
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (SITE_DIR / "index.html").write_text(render_index_html(index_posts), encoding="utf-8")
    (SITE_DIR / "government-class-quiz.html").write_text(render_quiz_html(), encoding="utf-8")

    print(f"Built {len(index_posts)} posts into {SITE_DIR}")
    print(f"Loaded {len(character_defs)} character definitions.")


if __name__ == "__main__":
    build()

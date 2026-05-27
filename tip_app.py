#!/usr/bin/env python3
"""Test & Inspection Plan - Standalone Application  Rev 7"""

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, datetime, tempfile
from pathlib import Path

try:
    import win32com.client
    OUTLOOK_OK = True
except ImportError:
    OUTLOOK_OK = False

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm, mm

# ── Colour palette (matches Excel) ─────────────────────────────────────────
C_NAVY    = "#1F3864"
C_NAVY2   = "#2F5496"
C_AMBER   = "#FFD966"
C_DGOLD   = "#A67C00"
C_CREAM   = "#FFFDE7"
C_SKYBLUE = "#DEEAF6"
C_LTGREEN = "#EDF7EE"
C_DKGREEN = "#1E4620"
C_GREY    = "#BFBFBF"

# ── Dropdown data ───────────────────────────────────────────────────────────
VM_OPTIONS = ["", "CMS", "MANUAL", "VISUAL", "PROCESS CONTROL PLAN",
              "SPECIAL PROCESS CONTROL", "OTHER", "N/A"]

MSA_OPTIONS = ["",
    "FULL GAUGE R&R AND BIAS STUDY", "MEASUREMENT UNCERTAINTY ANALYSIS",
    "ATTRIBUTE AGREEMENT ANALYSIS",
    "READ-ACROSS FROM AN APPROVED MEASUREMENT STUDY REPORT",
    "READ-ACROSS FROM DIMENSIONAL MEASUREMENT EQUIPMENT GUIDE MXG011",
    "TYPE 1 WITH BIAS STUDY", "AUTOMATED MEASUREMENT SYSTEM AND BIAS STUDY",
    "APPROVED CONTROL PLAN", "MSA NOT REQUIRED", "SPECIAL PROCESS CONTROL", "N/A"]

ET_OPTIONS = ["",
    "3D STRUCTURED LASER SCANNING","ACCELEROMETER","AIR GAUGE","AIRFLOW",
    "ANGLE PLATE","ARTICULATED ARM (eg. FARO / ROMOR)","AUTOMATIC HEIGHT GAUGE",
    "BINOCULARS","BORE MICROMETER  2 - POINT","BORE MICROMETER  3 - POINT",
    "BORE INDICATOR  2 - POINT","BORE INDICATOR  3 - POINT","BOROSCOPE",
    "CALIPER - DEPTH","CALIPER - DIAL","CALIPER - DIGITAL","CALIPER - VERNIER",
    "CALIPER (DIGITAL) & SCREW THREAD PLUG GAUGE","CALIPER / VISUAL",
    "CHAMFER GAUGE","CIRC TAPE","CMM",
    "CNC MACHINE  - IN CYCLE PROBING (CONTROL PLAN)",
    "COMPARATOR GAUGE","COMPARATOR GAUGE WITH DTI","COUNTERSINK GAUGE",
    "DEPTH GAUGE","DIATEST","DIGITAL HEIGHT GAUGE","DIGITAL PROTRACTOR",
    "DTI (CLOCK)","ELECTRICAL CONDUCTIVITY METER","ELECTRICAL CONTINUITY TESTER",
    "ENGINEERS SQUARE","FEELER GAUGE","GAUGE BLOCK & FEELER GAUGE","GAUGE BLOCKS",
    "GRATICULE","HAND HELD LASER SCANNER","HARDNESS TESTER","HELI-COIL DEPTH",
    "LASER TRACKER","MAGNIFICATION DEVICE","MAGNETIC PARTICLE INSPECTION",
    "MANUAL CMM","MICROMETER - ANALOGUE","MICROMETER - DIGITAL",
    "MICROMETER - INTERNAL","MICROMETER - DEEP THROAT","MICROMETER - DEPTH",
    "MICROMETER - POINTED MIC","MICROSCOPE","MODULAR BEAM COMPARATOR",
    "MOMENT WEIGHING SCALES","MOULD AND SHADOWGRAPH",
    "MOULD, SHADOWGRAPH & GRATICULE","OMNIGAUGE COMPARITOR","OTHER",
    "OPTICAL 3D MEASUREMENT DEVICE (eg ALICONA)","TOOLSETTER (eg PARLEC)",
    "PLAIN PLUG GAUGE","PRESSURE GAUGE","PRE-SWAGE INSERTS","PROFILE PROJECTOR",
    "PROTRACTOR (HELIX - GRATICULE)","PROTRACTOR (VERNIER)","RADIOGRAPHY",
    "RADIUS GAUGE","ROLLS-ROYCE APPROVED TOOLING/GAUGING","ROTARY TABLE WITH DTI",
    "ROUNDNESS & FORM MACHINES","ROUND UP TABLE","RULER",
    "SCREW THREAD PLUG GAUGES","SECTION CALIPER","SETTING BAR","SETTING RING",
    "SHADOWGRAPH","SNAP GAUGE","SPECIAL TO PRODUCT - DEPTH GAUGE",
    "SPECIAL TO PRODUCT GAUGE","SPIRIT LEVEL","SQUARE GAUGE",
    "STRAIGHT EDGE & DEPTH GAUGE",
    "SURFACE FINISH COMPARATOR (eg.RUBERT BLOCKS)","SURFACE ROUGHNESS METER",
    "TAPE MEASURE","THERMOMETER","THICKNESS METER","THREAD GAUGES",
    "THREAD PLUG GAUGES","TORQUE WRENCH","ULTRASOUND THICKNESS GAUGE",
    "UNIVERSAL BEVEL PROTRACTOR","VISUAL VERIFICATION","WEIGHING SCALES",
    "WIRE GAUGES","N/A"]

# ── Data model ──────────────────────────────────────────────────────────────
def new_card(n):
    return dict(item_num=n, op_no="", sampling="", meas_type="", spec="",
                units="", bubble="", equipment_name="", equipment_no="",
                description="", ref_loc="", process_ref="",
                verification_method="", equipment_type="", msa_method="",
                msa_ref="", responsible_person="", control_method="",
                reaction_plan="", comments="")

def new_plan():
    return dict(mva_author="", mva_number="",
                date=datetime.date.today().strftime("%d/%m/%Y"),
                mode="manual", tar_email="")

# ── Persistence ─────────────────────────────────────────────────────────────
def save_plan(path, plan, cards):
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"version": "1.0", "plan": plan, "cards": cards}, f, indent=2)

def load_plan(path):
    with open(path, "r", encoding="utf-8") as f:
        d = json.load(f)
    return d.get("plan", new_plan()), d.get("cards", [])

# ── NetInspect import ───────────────────────────────────────────────────────
NI = dict(op=4, sampling=7, desc=8, type=9, spec=10,
          units=11, bubble=12, proc=15, ref=18)

def import_netinspect(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    hdr = str(ws.cell(2, NI["op"]).value or "")
    if "operation" not in hdr.lower():
        raise ValueError("Not a standard NetInspect export (expected 'Operation' in col D, row 2).")
    cards = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) < max(NI.values()):
            continue
        op = str(row[NI["op"]-1] or "").strip()
        if not op:
            continue
        c = new_card(len(cards) + 1)
        c["op_no"]       = op
        c["sampling"]    = str(row[NI["sampling"]-1] or "")
        c["description"] = str(row[NI["desc"]-1]     or "")
        c["meas_type"]   = str(row[NI["type"]-1]     or "")
        c["spec"]        = str(row[NI["spec"]-1]      or "")
        c["units"]       = str(row[NI["units"]-1]     or "")
        c["bubble"]      = str(row[NI["bubble"]-1]    or "")
        c["ref_loc"]     = str(row[NI["ref"]-1]       or "")
        c["process_ref"] = str(row[NI["proc"]-1]      or "")
        cards.append(c)
    wb.close()
    return cards

# ── PDF export ──────────────────────────────────────────────────────────────
def _rl_color(h):
    r,g,b = int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
    return colors.Color(r/255,g/255,b/255)

RL_NAVY   = _rl_color("1F3864"); RL_GREEN  = _rl_color("1E4620")
RL_AMBER  = _rl_color("FFD966"); RL_GOLD   = _rl_color("A67C00")
RL_CREAM  = _rl_color("FFFDE7"); RL_SKY    = _rl_color("DEEAF6")
RL_LTGRN  = _rl_color("EDF7EE"); RL_GREY   = _rl_color("BFBFBF")

def _ps(name, size=8, bold=False, color=colors.black):
    fn = "Helvetica-Bold" if bold else "Helvetica"
    return ParagraphStyle(name, fontName=fn, fontSize=size,
                          textColor=color, leading=size+3)

PS_WH9B  = _ps("wh9b",  9, True,  colors.white)
PS_WH8B  = _ps("wh8b",  8, True,  colors.white)
PS_NV8B  = _ps("nv8b",  8, True,  RL_NAVY)
PS_GR8B  = _ps("gr8b",  8, True,  RL_GREEN)
PS_NV8   = _ps("nv8",   8, False, RL_NAVY)
PS_BK9   = _ps("bk9",   9, False, colors.black)

def _p(txt, style): return Paragraph(str(txt) if txt else "", style)

def _grid(rows_data, col_w, bg_pairs=None, vpad=3, hpad=4):
    ts = [("GRID",(0,0),(-1,-1),0.4,RL_GREY),
          ("LEFTPADDING",(0,0),(-1,-1),hpad),
          ("RIGHTPADDING",(0,0),(-1,-1),hpad),
          ("TOPPADDING",(0,0),(-1,-1),vpad),
          ("BOTTOMPADDING",(0,0),(-1,-1),vpad)]
    for (r,c1,c2),bg in (bg_pairs or []):
        ts.append(("BACKGROUND",(c1,r),(c2,r),bg))
    return Table(rows_data, colWidths=col_w, style=TableStyle(ts))

def export_pdf(filepath, plan, cards):
    doc = SimpleDocTemplate(filepath, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    W  = doc.width
    story = []

    # Title banner
    story.append(Table([[_p("TEST &amp; INSPECTION PLAN  -  Rev 7", PS_WH9B)]],
        colWidths=[W], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),RL_NAVY),
            ("LEFTPADDING",(0,0),(-1,-1),10),("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8)])))
    story.append(Spacer(1,4*mm))

    # Plan info
    story.append(_grid(
        [[_p("MVA Number:",PS_NV8B),_p(plan.get("mva_number",""),PS_BK9),
          _p("Prepared By:",PS_NV8B),_p(plan.get("mva_author",""),PS_BK9)],
         [_p("Date:",PS_NV8B),_p(plan.get("date",""),PS_BK9),
          _p("Input Method:",PS_NV8B),_p(plan.get("mode","").capitalize(),PS_BK9)]],
        [W*.15,W*.35,W*.15,W*.35],
        [((0,0,3),RL_SKY),((1,0,3),RL_SKY)]))
    story.append(Spacer(1,6*mm))

    mode = plan.get("mode","manual")
    for card in cards:
        story.extend(_card_story(card, mode, W))
        story.append(Spacer(1,3*mm))

    # Approval
    story.append(Spacer(1,5*mm))
    story.append(Table([[_p("AUTHORISATION &amp; APPROVAL",PS_WH9B)]],
        colWidths=[W],style=TableStyle([("BACKGROUND",(0,0),(-1,-1),RL_NAVY),
            ("LEFTPADDING",(0,0),(-1,-1),10),("TOPPADDING",(0,0),(-1,-1),6),
            ("BOTTOMPADDING",(0,0),(-1,-1),6)])))
    story.append(_grid(
        [[_p("Role",PS_NV8B),_p("Full Name",PS_NV8B),
          _p("Signature",PS_NV8B),_p("Date",PS_NV8B),_p("Emp No.",PS_NV8B)],
         [_p("Approved By (TAR Holder)",PS_BK9),"","","",""]],
        [W*.30,W*.20,W*.20,W*.15,W*.15],
        [((0,0,4),RL_SKY)], vpad=14))

    doc.build(story)

def _card_story(card, mode, W):
    items = []
    hdr_txt = (f"  ITEM {card['item_num']}"
               + (f"   Op: {card['op_no']}" if card.get("op_no") else "")
               + (f"   Sampling: {card['sampling']}" if card.get("sampling") else "")
               + (f"   Type: {card['meas_type']}" if card.get("meas_type") else "")
               + (f"   Spec: {card['spec']}{' '+card['units'] if card.get('units') else ''}"
                  if card.get("spec") else "")
               + (f"   Bubble #: {card['bubble']}" if card.get("bubble") else ""))

    rows = []
    row_bg = []

    if mode == "automated":
        rows.append([_p(hdr_txt, PS_WH9B)])
        rows.append([_p(card.get("description",""), PS_BK9)])
        row_bg = [((0,0,0),RL_NAVY),((1,0,0),RL_CREAM)]
    else:
        rows.append([_p(f"  ITEM {card['item_num']}", PS_WH9B)])
        rows.append([_grid(
            [[_p("Op No.",PS_NV8B),_p("Sampling",PS_NV8B),_p("Type",PS_NV8B),
              _p("Spec",PS_NV8B),_p("Units",PS_NV8B),_p("Bubble #",PS_NV8B)],
             [_p(card.get("op_no",""),PS_BK9),_p(card.get("sampling",""),PS_BK9),
              _p(card.get("meas_type",""),PS_BK9),_p(card.get("spec",""),PS_BK9),
              _p(card.get("units",""),PS_BK9),_p(card.get("bubble",""),PS_BK9)]],
            [W/6]*6, [((0,0,5),RL_AMBER)])])
        rows.append([_grid(
            [[_p("  Equipment Name:",PS_GR8B),_p(card.get("equipment_name",""),PS_BK9),
              _p("  Equipment No.:",PS_GR8B),_p(card.get("equipment_no",""),PS_BK9)]],
            [W*.25]*4,[((0,0,3),RL_LTGRN)])])
        rows.append([_p("  INSPECTION REQUIREMENT / DESCRIPTION:", PS_WH8B)])
        rows.append([_p(card.get("description",""), PS_BK9)])
        row_bg = [((0,0,0),RL_NAVY),((3,0,0),RL_GOLD),((4,0,0),RL_CREAM)]

    # Ref row (both modes)
    rows.append([_grid(
        [[_p(f"  Ref: {card.get('ref_loc','')}", PS_NV8),
          _p(f"  Process Ref: {card.get('process_ref','')}", PS_NV8)]],
        [W*.5,W*.5],[((0,0,1),RL_SKY)])])

    # TAR header
    rows.append([_p("  TAR HOLDER - PLEASE COMPLETE:", PS_WH8B)])
    row_bg.append(((len(rows)-1,0,0),RL_GREEN))

    # TAR data
    rows.append([_grid(
        [[_p("  Verification Method:",PS_GR8B),_p(card.get("verification_method",""),PS_BK9),
          _p("  Equipment Type:",PS_GR8B),_p(card.get("equipment_type",""),PS_BK9)],
         [_p("  MSA Verification Method:",PS_GR8B),_p(card.get("msa_method",""),PS_BK9),
          _p("  MSA Ref / Justification:",PS_GR8B),_p(card.get("msa_ref",""),PS_BK9)],
         [_p("  Responsible Person:",PS_GR8B),_p(card.get("responsible_person",""),PS_BK9),
          _p("  Control Method:",PS_GR8B),_p(card.get("control_method",""),PS_BK9)],
         [_p("  Reaction Plan:",PS_GR8B),_p(card.get("reaction_plan",""),PS_BK9),
          _p("  Comments:",PS_GR8B),_p(card.get("comments",""),PS_BK9)]],
        [W*.25]*4,
        [((r,0,0),RL_LTGRN) for r in range(4)]+[((r,2,2),RL_LTGRN) for r in range(4)])])

    ts = [("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
          ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
          ("BOX",(0,0),(-1,-1),1.5,RL_NAVY)]
    for (ri,c1,c2),bg in row_bg:
        ts.append(("BACKGROUND",(c1,ri),(c2,ri),bg))

    items.append(Table(rows, colWidths=[W], style=TableStyle(ts)))
    return items

# ── Email ───────────────────────────────────────────────────────────────────
def send_outlook(to_addr, subject, body, attachments=None):
    if not OUTLOOK_OK:
        raise RuntimeError("pywin32 not installed")
    ol = win32com.client.Dispatch("Outlook.Application")
    m = ol.CreateItem(0)
    m.To, m.Subject, m.Body = to_addr, subject, body
    for a in (attachments or []):
        if os.path.exists(a):
            m.Attachments.Add(a)
    m.Display()

# ── Card edit dialog ────────────────────────────────────────────────────────
class CardDialog(ctk.CTkToplevel):
    def __init__(self, parent, card, mode, on_save):
        super().__init__(parent)
        self.card    = dict(card)
        self.mode    = mode
        self.on_save = on_save
        self.title(f"ITEM {card['item_num']}")
        self.geometry("760x660")
        self.resizable(True, True)
        self.grab_set()
        self._vars = {}
        self._build()

    def _build(self):
        scr = ctk.CTkScrollableFrame(self)
        scr.pack(fill="both", expand=True, padx=6, pady=6)
        scr.columnconfigure([1,3], weight=1)

        # Title
        hf = ctk.CTkFrame(scr, fg_color=C_NAVY, corner_radius=4)
        hf.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0,8))
        ctk.CTkLabel(hf, text=f"  ITEM {self.card['item_num']}",
                     text_color="white",
                     font=ctk.CTkFont("Arial",11,"bold")).pack(side="left",pady=5)

        r = 1
        if self.mode == "manual":
            for lbl, key, col_offset in [
                ("Op No.:", "op_no", 0), ("Sampling:", "sampling", 2),
                ("Type:", "meas_type", 0), ("Spec:", "spec", 2),
                ("Units:", "units", 0), ("Bubble #:", "bubble", 2)
            ]:
                if col_offset == 0:
                    ro = r
                ctk.CTkLabel(scr, text=lbl, anchor="e").grid(
                    row=ro, column=col_offset, padx=(4,2), pady=2, sticky="e")
                v = ctk.StringVar(value=self.card.get(key,""))
                self._vars[key] = v
                ctk.CTkEntry(scr, textvariable=v).grid(
                    row=ro, column=col_offset+1, padx=(0,8), pady=2, sticky="ew")
                if col_offset == 2:
                    r += 1

        # Equipment (green bg)
        ef = ctk.CTkFrame(scr, fg_color=C_LTGREEN, corner_radius=0)
        ef.grid(row=r, column=0, columnspan=4, sticky="ew", pady=(8,1))
        ef.columnconfigure([1,3], weight=1)
        ctk.CTkLabel(ef, text="Equipment Name:", anchor="e").grid(
            row=0,column=0,padx=(8,2),pady=4,sticky="e")
        v = ctk.StringVar(value=self.card.get("equipment_name",""))
        self._vars["equipment_name"] = v
        ctk.CTkEntry(ef, textvariable=v).grid(row=0,column=1,padx=(0,8),pady=4,sticky="ew")
        ctk.CTkLabel(ef, text="Equipment No.:", anchor="e").grid(
            row=0,column=2,padx=(8,2),pady=4,sticky="e")
        v2 = ctk.StringVar(value=self.card.get("equipment_no",""))
        self._vars["equipment_no"] = v2
        ctk.CTkEntry(ef, textvariable=v2, width=140).grid(
            row=0,column=3,padx=(0,8),pady=4,sticky="ew")
        r += 1

        # Description header
        dh = ctk.CTkFrame(scr, fg_color=C_DGOLD, corner_radius=0)
        dh.grid(row=r, column=0, columnspan=4, sticky="ew", pady=(6,0))
        ctk.CTkLabel(dh, text="  INSPECTION REQUIREMENT / DESCRIPTION:",
                     text_color="white",
                     font=ctk.CTkFont("Arial",8,"bold")).pack(side="left",pady=3)
        r += 1
        self._desc = ctk.CTkTextbox(scr, height=80, fg_color=C_CREAM,
                                     border_width=1, border_color=C_GREY)
        self._desc.grid(row=r, column=0, columnspan=4, sticky="ew", pady=(0,6))
        self._desc.insert("1.0", self.card.get("description",""))
        r += 1

        # Ref row
        rf = ctk.CTkFrame(scr, fg_color=C_SKYBLUE, corner_radius=0)
        rf.grid(row=r, column=0, columnspan=4, sticky="ew", pady=1)
        rf.columnconfigure([1,3], weight=1)
        ctk.CTkLabel(rf, text="Ref Location:", anchor="e").grid(
            row=0,column=0,padx=(8,2),pady=4,sticky="e")
        vr = ctk.StringVar(value=self.card.get("ref_loc",""))
        self._vars["ref_loc"] = vr
        ctk.CTkEntry(rf, textvariable=vr).grid(row=0,column=1,padx=(0,8),pady=4,sticky="ew")
        ctk.CTkLabel(rf, text="Process Ref:", anchor="e").grid(
            row=0,column=2,padx=(8,2),pady=4,sticky="e")
        vp = ctk.StringVar(value=self.card.get("process_ref",""))
        self._vars["process_ref"] = vp
        ctk.CTkEntry(rf, textvariable=vp).grid(row=0,column=3,padx=(0,8),pady=4,sticky="ew")
        r += 1

        # TAR header
        th = ctk.CTkFrame(scr, fg_color=C_DKGREEN, corner_radius=0)
        th.grid(row=r, column=0, columnspan=4, sticky="ew", pady=(10,0))
        ctk.CTkLabel(th, text="  TAR HOLDER - PLEASE COMPLETE:",
                     text_color="white",
                     font=ctk.CTkFont("Arial",8,"bold")).pack(side="left",pady=3)
        r += 1

        tf = ctk.CTkFrame(scr, fg_color=C_LTGREEN, corner_radius=0)
        tf.grid(row=r, column=0, columnspan=4, sticky="ew", pady=1)
        tf.columnconfigure([1,3], weight=1)

        tar_fields = [
            ("Verification Method:", "verification_method", "dropdown", VM_OPTIONS),
            ("Equipment Type:", "equipment_type", "dropdown", ET_OPTIONS),
            ("MSA Verification Method:", "msa_method", "dropdown", MSA_OPTIONS),
            ("MSA Ref / Justification:", "msa_ref", "entry", None),
            ("Responsible Person:", "responsible_person", "entry", None),
            ("Control Method:", "control_method", "entry", None),
            ("Reaction Plan:", "reaction_plan", "entry", None),
            ("Comments:", "comments", "entry", None),
        ]
        for i, (lbl, key, wtype, opts) in enumerate(tar_fields):
            row_idx, col_idx = divmod(i, 2)
            ctk.CTkLabel(tf, text=lbl, anchor="e",
                         font=ctk.CTkFont("Arial",9,"bold"),
                         text_color=C_DKGREEN).grid(
                             row=row_idx, column=col_idx*2,
                             padx=(8,2), pady=3, sticky="e")
            v = ctk.StringVar(value=self.card.get(key,""))
            self._vars[key] = v
            if wtype == "dropdown":
                ctk.CTkOptionMenu(tf, variable=v, values=opts, width=180).grid(
                    row=row_idx, column=col_idx*2+1,
                    padx=(0,8), pady=3, sticky="ew")
            else:
                ctk.CTkEntry(tf, textvariable=v).grid(
                    row=row_idx, column=col_idx*2+1,
                    padx=(0,8), pady=3, sticky="ew")
        r += 1

        # Buttons
        bf = ctk.CTkFrame(scr, fg_color="transparent")
        bf.grid(row=r, column=0, columnspan=4, pady=12)
        ctk.CTkButton(bf, text="Save", width=120, fg_color=C_NAVY,
                      command=self._save).pack(side="left", padx=10)
        ctk.CTkButton(bf, text="Cancel", width=100, fg_color="gray",
                      command=self.destroy).pack(side="left", padx=4)

    def _save(self):
        for key, var in self._vars.items():
            self.card[key] = var.get()
        self.card["description"] = self._desc.get("1.0","end-1c")
        self.on_save(self.card)
        self.destroy()


# ── Main Application ────────────────────────────────────────────────────────
class TIPApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title("Test & Inspection Plan  |  Rev 7")
        self.geometry("1120x760")
        self.minsize(900, 600)

        self.plan      = new_plan()
        self.cards     = []
        self._tip_path = None
        self._setup_vars = {}
        self._mode_var = ctk.StringVar(value="manual")

        self._build_header()
        self._build_tabs()
        self._build_status()

    # ── Header ─────────────────────────────────────────────────────────────
    def _build_header(self):
        h = ctk.CTkFrame(self, fg_color=C_NAVY, height=52, corner_radius=0)
        h.pack(fill="x")
        h.pack_propagate(False)
        ctk.CTkLabel(h, text="  TEST & INSPECTION PLAN",
                     text_color="white",
                     font=ctk.CTkFont("Arial",16,"bold")).pack(side="left",padx=10)
        ctk.CTkLabel(h, text="Rev 7",
                     text_color=C_AMBER,
                     font=ctk.CTkFont("Arial",12)).pack(side="left")

    # ── Status bar ──────────────────────────────────────────────────────────
    def _build_status(self):
        self._sv = tk.StringVar(value="Ready")
        sb = ctk.CTkFrame(self, fg_color="#E0E0E0", height=24, corner_radius=0)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        ctk.CTkLabel(sb, textvariable=self._sv,
                     font=ctk.CTkFont("Arial",10),
                     text_color="#333333").pack(side="left",padx=10)

    def status(self, msg): self._sv.set(msg)

    # ── Tabs ────────────────────────────────────────────────────────────────
    def _build_tabs(self):
        self.nb = ctk.CTkTabview(self, corner_radius=4)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        for t in ["Plan Setup", "Inspection Data", "Send to TAR Holder", "TAR Approval"]:
            self.nb.add(t)
        self._build_setup(self.nb.tab("Plan Setup"))
        self._build_data(self.nb.tab("Inspection Data"))
        self._build_send(self.nb.tab("Send to TAR Holder"))
        self._build_approval(self.nb.tab("TAR Approval"))

    # ══════════════════════════════════════════════════════════════════════
    # TAB 1 – PLAN SETUP
    # ══════════════════════════════════════════════════════════════════════
    def _build_setup(self, parent):
        frm = ctk.CTkScrollableFrame(parent)
        frm.pack(fill="both", expand=True, padx=4, pady=4)
        frm.columnconfigure(1, weight=1)

        def add_row(label, key, r, date_btn=False):
            ctk.CTkLabel(frm, text=label, anchor="e", width=180).grid(
                row=r, column=0, padx=(4,6), pady=7, sticky="e")
            v = ctk.StringVar(value=self.plan.get(key,""))
            self._setup_vars[key] = v
            ctk.CTkEntry(frm, textvariable=v, width=380).grid(
                row=r, column=1, sticky="ew", padx=(0,4), pady=7)
            if date_btn:
                ctk.CTkButton(frm, text="Today", width=70,
                              command=lambda vv=v: vv.set(
                                  datetime.date.today().strftime("%d/%m/%Y"))
                              ).grid(row=r, column=2, padx=4, pady=7)

        add_row("MVA Author (Prepared By):", "mva_author", 0)
        add_row("MVA Number:", "mva_number", 1)
        add_row("Date (DD/MM/YYYY):", "date", 2, date_btn=True)

        ctk.CTkLabel(frm, text="Input Method:", anchor="e", width=180).grid(
            row=3, column=0, padx=(4,6), pady=7, sticky="e")
        mf = ctk.CTkFrame(frm, fg_color="transparent")
        mf.grid(row=3, column=1, sticky="w", pady=7)
        ctk.CTkRadioButton(mf, text="Manual", variable=self._mode_var,
                           value="manual",
                           command=self._mode_changed).pack(side="left",padx=14)
        ctk.CTkRadioButton(mf, text="Automated (import NetInspect)",
                           variable=self._mode_var, value="automated",
                           command=self._mode_changed).pack(side="left",padx=14)

        add_row("TAR Holder Email:", "tar_email", 4)

        sep = ctk.CTkFrame(frm, fg_color=C_GREY, height=1)
        sep.grid(row=5, column=0, columnspan=3, sticky="ew", padx=4, pady=12)

        bf = ctk.CTkFrame(frm, fg_color="transparent")
        bf.grid(row=6, column=0, columnspan=3, pady=4)
        ctk.CTkButton(bf, text="New Plan", width=130, fg_color="#666666",
                      command=self._new_plan).pack(side="left",padx=8)
        ctk.CTkButton(bf, text="Load Plan", width=130, fg_color=C_NAVY2,
                      command=self._load_plan).pack(side="left",padx=8)
        ctk.CTkButton(bf, text="Save Plan", width=130, fg_color=C_NAVY,
                      command=self._save_plan).pack(side="left",padx=8)

    def _apply_setup(self):
        for k, v in self._setup_vars.items():
            self.plan[k] = v.get()
        self.plan["mode"] = self._mode_var.get()

    def _mode_changed(self):
        self._apply_setup()
        self._refresh_data()

    def _new_plan(self):
        if self.cards and not messagebox.askyesno(
                "New Plan","Discard current plan and start fresh?"):
            return
        self.plan, self.cards, self._tip_path = new_plan(), [], None
        for k, v in self._setup_vars.items():
            v.set(self.plan.get(k,""))
        self._mode_var.set("manual")
        self._refresh_data()
        self.status("New blank plan created.")

    def _save_plan(self):
        self._apply_setup()
        path = self._tip_path or filedialog.asksaveasfilename(
            defaultextension=".tip",
            filetypes=[("TIP Plan","*.tip"),("All Files","*.*")],
            title="Save Plan")
        if not path: return
        self._tip_path = path
        save_plan(path, self.plan, self.cards)
        self.status(f"Saved: {path}")

    def _load_plan(self):
        path = filedialog.askopenfilename(
            filetypes=[("TIP Plan","*.tip"),("All Files","*.*")],
            title="Load Plan")
        if not path: return
        self.plan, self.cards = load_plan(path)
        self._tip_path = path
        for k, v in self._setup_vars.items():
            v.set(self.plan.get(k,""))
        self._mode_var.set(self.plan.get("mode","manual"))
        self._refresh_data()
        self.status(f"Loaded: {path}  ({len(self.cards)} cards)")

    # ══════════════════════════════════════════════════════════════════════
    # TAB 2 – INSPECTION DATA
    # ══════════════════════════════════════════════════════════════════════
    def _build_data(self, parent):
        self._data_parent  = parent
        self._data_content = None
        self._refresh_data()

    def _refresh_data(self):
        if self._data_content:
            self._data_content.destroy()

        p = self._data_parent
        self._data_content = ctk.CTkFrame(p, fg_color="transparent")
        self._data_content.pack(fill="both", expand=True)

        mode = self.plan.get("mode","manual")

        # Toolbar
        tb = ctk.CTkFrame(self._data_content, fg_color="#E8EEF7",
                           height=44, corner_radius=4)
        tb.pack(fill="x", padx=4, pady=(4,2))
        tb.pack_propagate(False)

        self._count_lbl = ctk.CTkLabel(tb,
            text=f"{len(self.cards)} cards",
            font=ctk.CTkFont("Arial",11), text_color=C_NAVY)
        self._count_lbl.pack(side="right", padx=14)

        if mode == "automated":
            ctk.CTkButton(tb, text="Import NetInspect File", width=200,
                          fg_color=C_NAVY,
                          command=self._import_ni).pack(side="left",padx=8,pady=6)
        else:
            for txt, cmd, clr in [
                ("+ Add Card",      self._add_card,    C_NAVY),
                ("Edit Selected",   self._edit_card,   C_NAVY2),
                ("Delete Selected", self._del_card,    "#8B0000"),
                ("Clear All",       self._clear_cards, "#666666"),
            ]:
                ctk.CTkButton(tb, text=txt, width=130, fg_color=clr,
                              command=cmd).pack(side="left",padx=4,pady=6)

        # Treeview
        tf = ctk.CTkFrame(self._data_content, fg_color="white")
        tf.pack(fill="both", expand=True, padx=4, pady=(2,4))

        cols = ("#","Op No.","Sampling","Type","Spec / Units","Description")
        self._tree = ttk.Treeview(tf, columns=cols, show="headings",
                                   selectmode="browse")
        cw = [38,72,80,80,130,0]
        for col, w in zip(cols, cw):
            self._tree.heading(col, text=col)
            self._tree.column(col, width=w, stretch=(w==0))

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._tree.bind("<Double-1>", lambda e: self._edit_card())

        s = ttk.Style()
        s.configure("Treeview", font=("Arial",9), rowheight=22)
        s.configure("Treeview.Heading", font=("Arial",9,"bold"))

        self._rebuild_tree()

    def _rebuild_tree(self):
        self._tree.delete(*self._tree.get_children())
        for c in self.cards:
            self._tree.insert("","end", values=(
                c["item_num"], c.get("op_no",""), c.get("sampling",""),
                c.get("meas_type",""),
                f"{c.get('spec','')} {c.get('units','')}".strip(),
                (c.get("description") or "")[:70]))
        n = len(self.cards)
        self._count_lbl.configure(text=f"{n} card{'s' if n!=1 else ''}")

    def _import_ni(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel","*.xlsx *.xls *.xlsm"),("All Files","*.*")],
            title="Select NetInspect Export File")
        if not path: return
        try:
            self.cards = import_netinspect(path)
            self._rebuild_tree()
            self.status(f"Imported {len(self.cards)} items from {Path(path).name}")
        except Exception as ex:
            messagebox.showerror("Import Error", str(ex))

    def _add_card(self):
        card = new_card(len(self.cards)+1)
        def on_save(c):
            self.cards.append(c)
            self._rebuild_tree()
            self.status(f"Card {c['item_num']} added.")
        CardDialog(self, card, self.plan.get("mode","manual"), on_save)

    def _edit_card(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("Edit","Select a card first."); return
        idx = self._tree.index(sel[0])
        def on_save(c):
            self.cards[idx] = c
            self._rebuild_tree()
            self.status(f"Card {c['item_num']} updated.")
        CardDialog(self, self.cards[idx], self.plan.get("mode","manual"), on_save)

    def _del_card(self):
        sel = self._tree.selection()
        if not sel: return
        idx = self._tree.index(sel[0])
        if messagebox.askyesno("Delete", f"Delete ITEM {self.cards[idx]['item_num']}?"):
            del self.cards[idx]
            for i, c in enumerate(self.cards, 1):
                c["item_num"] = i
            self._rebuild_tree()

    def _clear_cards(self):
        if self.cards and messagebox.askyesno("Clear All","Remove all cards?"):
            self.cards = []
            self._rebuild_tree()

    # ══════════════════════════════════════════════════════════════════════
    # TAB 3 – SEND TO TAR HOLDER
    # ══════════════════════════════════════════════════════════════════════
    def _build_send(self, parent):
        frm = ctk.CTkScrollableFrame(parent)
        frm.pack(fill="both", expand=True, padx=4, pady=4)
        frm.columnconfigure(1, weight=1)

        ctk.CTkLabel(frm, text="PLAN SUMMARY",
                     font=ctk.CTkFont("Arial",11,"bold"),
                     text_color=C_NAVY).grid(row=0,column=0,columnspan=3,
                                              sticky="w",padx=8,pady=(8,2))
        self._sum_lbl = ctk.CTkLabel(frm, text="", anchor="w", justify="left",
                                      font=ctk.CTkFont("Arial",10))
        self._sum_lbl.grid(row=1,column=0,columnspan=3,sticky="w",padx=24,pady=(0,12))

        ctk.CTkButton(frm, text="Refresh Summary", width=150, fg_color="#666666",
                      command=self._refresh_summary
                      ).grid(row=2,column=0,columnspan=3,sticky="w",padx=8,pady=(0,8))

        ctk.CTkLabel(frm, text="TAR Holder Email:", anchor="e", width=170).grid(
            row=3,column=0,padx=(4,6),pady=7,sticky="e")
        self._email_var = ctk.StringVar()
        ctk.CTkEntry(frm, textvariable=self._email_var, width=380).grid(
            row=3,column=1,sticky="ew",padx=(0,4),pady=7)

        sep = ctk.CTkFrame(frm, fg_color=C_GREY, height=1)
        sep.grid(row=4,column=0,columnspan=3,sticky="ew",padx=4,pady=12)

        bf = ctk.CTkFrame(frm, fg_color="transparent")
        bf.grid(row=5,column=0,columnspan=3,pady=4)
        ctk.CTkButton(bf, text="Send Email via Outlook", width=200,
                      fg_color=C_NAVY,
                      command=self._send_email).pack(side="left",padx=8)
        ctk.CTkButton(bf, text="Export Draft PDF", width=160,
                      fg_color=C_NAVY2,
                      command=self._export_pdf).pack(side="left",padx=8)
        ctk.CTkButton(bf, text="Save Plan (.tip)", width=150,
                      fg_color="#666666",
                      command=self._save_plan).pack(side="left",padx=8)

    def _refresh_summary(self):
        self._apply_setup()
        p = self.plan
        self._sum_lbl.configure(text=(
            f"MVA Number:   {p.get('mva_number','(not set)')}\n"
            f"Prepared By:  {p.get('mva_author','(not set)')}\n"
            f"Date:         {p.get('date','(not set)')}\n"
            f"Input Method: {p.get('mode','manual').capitalize()}\n"
            f"Cards:        {len(self.cards)}"))
        self._email_var.set(p.get("tar_email",""))

    def _send_email(self):
        self._apply_setup()
        to = self._email_var.get().strip()
        if not to:
            messagebox.showwarning("Send Email","Enter the TAR Holder email."); return
        if not self.plan.get("mva_number"):
            messagebox.showwarning("Send Email","Enter the MVA Number first."); return
        if not self.cards:
            messagebox.showwarning("Send Email","No inspection cards to send."); return

        tip_path = self._tip_path or os.path.join(
            tempfile.gettempdir(), f"TIP_{self.plan['mva_number']}.tip")
        save_plan(tip_path, self.plan, self.cards)

        pdf_path = os.path.join(tempfile.gettempdir(),
                                 f"TIP_{self.plan['mva_number']}_Draft.pdf")
        try: export_pdf(pdf_path, self.plan, self.cards)
        except Exception: pdf_path = None

        p = self.plan
        body = (
            f"Dear TAR Holder,\n\n"
            f"Please find attached the Test & Inspection Plan for review.\n\n"
            f"  MVA Number  : {p.get('mva_number','')}\n"
            f"  Date        : {p.get('date','')}\n"
            f"  Prepared By : {p.get('mva_author','')}\n\n"
            + ("All inspection characteristics were automatically imported from NetInspect.\n"
               if p.get("mode")=="automated" else
               "This plan was prepared using Manual Entry mode.\n")
            + f"\nACTION REQUIRED:\n"
              f"  1. Open the .tip file using the TIP Application\n"
              f"  2. Go to the 'TAR Approval' tab\n"
              f"  3. Complete all green TAR fields for each item\n"
              f"  4. Export the Approved PDF, sign it, and email it back\n\n"
              f"Kind regards,\n{p.get('mva_author','[Operator]')}"
        )
        subject = f"ACTION REQUIRED: Test & Inspection Plan - {p.get('mva_number','')}"
        attachments = [a for a in [tip_path, pdf_path] if a and os.path.exists(a)]

        if OUTLOOK_OK:
            try:
                send_outlook(to, subject, body, attachments)
                self.status("Email draft opened in Outlook.")
            except Exception as ex:
                messagebox.showerror("Email Error", str(ex))
        else:
            messagebox.showinfo("Email (Outlook not available)",
                f"To: {to}\nSubject: {subject}\n\n"
                f"Plan saved to:\n{tip_path}")

    def _export_pdf(self, approved=False):
        self._apply_setup()
        if not self.cards:
            messagebox.showwarning("Export PDF","No cards to export."); return
        suffix = "_Approved" if approved else "_Draft"
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF","*.pdf"),("All Files","*.*")],
            initialfile=f"TIP_{self.plan.get('mva_number','PLAN')}{suffix}.pdf",
            title="Save PDF")
        if not path: return
        try:
            export_pdf(path, self.plan, self.cards)
            self.status(f"PDF saved: {path}")
            if messagebox.askyesno("PDF Saved","Open the PDF now?"):
                os.startfile(path)
        except Exception as ex:
            messagebox.showerror("PDF Error", str(ex))

    # ══════════════════════════════════════════════════════════════════════
    # TAB 4 – TAR APPROVAL
    # ══════════════════════════════════════════════════════════════════════
    def _build_approval(self, parent):
        main = ctk.CTkFrame(parent, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=4, pady=4)

        # Toolbar
        tb = ctk.CTkFrame(main, fg_color="#E8EEF7", height=46, corner_radius=4)
        tb.pack(fill="x", pady=(0,4))
        tb.pack_propagate(False)

        ctk.CTkButton(tb, text="Load Plan (.tip)", width=160, fg_color=C_NAVY,
                      command=self._tar_load).pack(side="left",padx=8,pady=7)
        ctk.CTkButton(tb, text="Use Current Plan", width=155, fg_color=C_NAVY2,
                      command=self._tar_use_current).pack(side="left",padx=4,pady=7)
        self._tar_info = ctk.CTkLabel(tb, text="No plan loaded",
                                       text_color="#555555",
                                       font=ctk.CTkFont("Arial",10))
        self._tar_info.pack(side="left", padx=16)

        ctk.CTkButton(tb, text="Export Approved PDF", width=180,
                      fg_color=C_DKGREEN,
                      command=self._tar_export).pack(side="right",padx=8,pady=7)
        ctk.CTkButton(tb, text="Email Back to Author", width=170,
                      fg_color=C_DGOLD,
                      command=self._tar_email).pack(side="right",padx=4,pady=7)

        # Split: tree (top) + TAR edit panel (bottom)
        self._tar_tree_frame = ctk.CTkFrame(main, fg_color="white")
        self._tar_tree_frame.pack(fill="both", expand=True)

        cols = ("#","Op No.","Description","Verif. Method","Equip. Type","Responsible")
        self._tar_tree = ttk.Treeview(self._tar_tree_frame, columns=cols,
                                       show="headings", selectmode="browse", height=10)
        cw2 = [38,72,0,130,110,120]
        for col,w in zip(cols,cw2):
            self._tar_tree.heading(col, text=col)
            self._tar_tree.column(col, width=w, stretch=(w==0))
        vsb2 = ttk.Scrollbar(self._tar_tree_frame, orient="vertical",
                               command=self._tar_tree.yview)
        self._tar_tree.configure(yscrollcommand=vsb2.set)
        self._tar_tree.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="right", fill="y")
        self._tar_tree.bind("<<TreeviewSelect>>", self._tar_select)

        # TAR edit panel
        self._tar_panel = ctk.CTkFrame(main, fg_color=C_LTGREEN, height=190)
        self._tar_panel.pack(fill="x", pady=(4,0))
        self._tar_panel.pack_propagate(False)
        self._tar_placeholder = ctk.CTkLabel(
            self._tar_panel,
            text="Select a card above to edit its TAR fields",
            text_color="#666666")
        self._tar_placeholder.pack(expand=True)

        self._tar_plan  = None
        self._tar_cards = []
        self._tar_cur   = None
        self._tar_vars  = {}

    def _tar_populate_tree(self):
        self._tar_tree.delete(*self._tar_tree.get_children())
        for c in self._tar_cards:
            self._tar_tree.insert("","end", values=(
                c["item_num"], c.get("op_no",""),
                (c.get("description","") or "")[:55],
                c.get("verification_method","") or "-",
                c.get("equipment_type","") or "-",
                c.get("responsible_person","") or "-"))

    def _tar_load(self):
        path = filedialog.askopenfilename(
            filetypes=[("TIP Plan","*.tip"),("All Files","*.*")],
            title="Load Plan for TAR Review")
        if not path: return
        self._tar_plan, self._tar_cards = load_plan(path)
        self._tar_tip_path = path
        self._tar_populate_tree()
        self._tar_info.configure(
            text=f"MVA: {self._tar_plan.get('mva_number','')}  |  "
                 f"{len(self._tar_cards)} cards")
        self.status(f"TAR loaded: {path}")

    def _tar_use_current(self):
        self._apply_setup()
        self._tar_plan  = dict(self.plan)
        self._tar_cards = [dict(c) for c in self.cards]
        self._tar_tip_path = self._tip_path
        self._tar_populate_tree()
        self._tar_info.configure(
            text=f"MVA: {self._tar_plan.get('mva_number','')}  |  "
                 f"{len(self._tar_cards)} cards (current plan)")

    def _tar_save_current(self):
        if self._tar_cur is None: return
        c = self._tar_cards[self._tar_cur]
        for k, v in self._tar_vars.items():
            c[k] = v.get()

    def _tar_select(self, _event=None):
        sel = self._tar_tree.selection()
        if not sel: return
        self._tar_save_current()
        self._tar_cur = self._tar_tree.index(sel[0])
        self._tar_populate_tree()
        self._tar_show_panel(self._tar_cards[self._tar_cur])

    def _tar_show_panel(self, card):
        for w in self._tar_panel.winfo_children():
            w.destroy()
        self._tar_vars = {}

        hf = ctk.CTkFrame(self._tar_panel, fg_color=C_DKGREEN, height=26, corner_radius=0)
        hf.pack(fill="x")
        ctk.CTkLabel(hf, text=f"  TAR HOLDER  -  ITEM {card['item_num']}",
                     text_color="white",
                     font=ctk.CTkFont("Arial",9,"bold")).pack(side="left",pady=2)

        gf = ctk.CTkFrame(self._tar_panel, fg_color=C_LTGREEN)
        gf.pack(fill="both", expand=True, padx=4, pady=4)
        gf.columnconfigure([1,3], weight=1)

        fields = [
            ("Verification Method:", "verification_method", "dd", VM_OPTIONS),
            ("Equipment Type:",      "equipment_type",       "dd", ET_OPTIONS),
            ("MSA Method:",          "msa_method",           "dd", MSA_OPTIONS),
            ("MSA Ref:",             "msa_ref",               "e",  None),
            ("Responsible Person:",  "responsible_person",    "e",  None),
            ("Control Method:",      "control_method",        "e",  None),
            ("Reaction Plan:",       "reaction_plan",          "e",  None),
            ("Comments:",            "comments",              "e",  None),
        ]
        for i, (lbl, key, wt, opts) in enumerate(fields):
            row, col = divmod(i, 2)
            ctk.CTkLabel(gf, text=lbl, anchor="e",
                         font=ctk.CTkFont("Arial",9,"bold"),
                         text_color=C_DKGREEN).grid(
                             row=row, column=col*2, padx=(6,2), pady=2, sticky="e")
            v = ctk.StringVar(value=card.get(key,""))
            self._tar_vars[key] = v
            if wt == "dd":
                ctk.CTkOptionMenu(gf, variable=v, values=opts, width=175).grid(
                    row=row, column=col*2+1, padx=(0,8), pady=2, sticky="ew")
            else:
                ctk.CTkEntry(gf, textvariable=v).grid(
                    row=row, column=col*2+1, padx=(0,8), pady=2, sticky="ew")

    def _tar_export(self):
        self._tar_save_current()
        if not self._tar_cards:
            messagebox.showwarning("Export","No plan loaded in TAR Approval tab."); return
        plan = self._tar_plan or self.plan
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF","*.pdf"),("All Files","*.*")],
            initialfile=f"TIP_{plan.get('mva_number','PLAN')}_Approved.pdf",
            title="Save Approved PDF")
        if not path: return
        try:
            export_pdf(path, plan, self._tar_cards)
            if hasattr(self,"_tar_tip_path") and self._tar_tip_path:
                save_plan(self._tar_tip_path, plan, self._tar_cards)
            self.status(f"Approved PDF: {path}")
            if messagebox.askyesno("PDF Saved","Open the PDF now?"):
                os.startfile(path)
        except Exception as ex:
            messagebox.showerror("PDF Error", str(ex))

    def _tar_email(self):
        self._tar_save_current()
        if not self._tar_cards:
            messagebox.showwarning("Email","No plan loaded."); return
        plan  = self._tar_plan or self.plan
        mva   = plan.get("mva_number","PLAN")

        dlg   = ctk.CTkInputDialog(text="Author's email address:", title="Email Back to Author")
        to    = dlg.get_input()
        if not to: return

        pdf_path = os.path.join(tempfile.gettempdir(), f"TIP_{mva}_Approved.pdf")
        try: export_pdf(pdf_path, plan, self._tar_cards)
        except Exception as ex:
            messagebox.showerror("PDF Error",str(ex)); return

        body = (f"Dear {plan.get('mva_author','Author')},\n\n"
                f"Please find attached the approved Test & Inspection Plan.\n\n"
                f"  MVA Number  : {mva}\n\n"
                f"Kind regards,\n[TAR Holder]")
        subject = f"APPROVED: Test & Inspection Plan - {mva}"

        if OUTLOOK_OK:
            try:
                send_outlook(to, subject, body, [pdf_path])
                self.status("Email draft opened in Outlook.")
            except Exception as ex:
                messagebox.showerror("Email Error", str(ex))
        else:
            messagebox.showinfo("Email (Outlook not available)",
                f"To: {to}\nSubject: {subject}\n\nPDF: {pdf_path}")


# ── Entry point ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = TIPApp()
    app.mainloop()

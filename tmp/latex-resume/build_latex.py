from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
D=Document('output/resume/Moein-Razavi-Unified-Resume.docx')
replacements={
'deterministic validation with up to 3 total attempts':'deterministic validation and bounded corrective retries',
'Designed LangChain/GPT-4o routing':'Designed the routing and answer-review architecture using LangChain/GPT-4o',
', with conversational context and domain-rule review':', with conversational context and domain-rule checks',
'Selected PatchTST using Python/PyTorch and Optuna on Vertex AI GPUs.':'Selected PatchTST through model comparisons; tuned candidates with Python/PyTorch and Optuna on Vertex AI GPUs.',
'Migrated Clinical from UIKit to MVVM-Coordinator':'Implemented Clinical\'s migration from UIKit to an MVVM-Coordinator architecture',
'Clink Social: Built location-based discovery and messaging with SwiftUI, MapKit, and Firebase Auth/Firestore/FCM. Implemented server-side feed aggregation and notifications in TypeScript Cloud Functions, plus QR-based cross-device sign-in with time-limited OAuth/JWT tokens.':'Clink Social: Built SwiftUI/MapKit discovery and messaging with Firebase. Implemented TypeScript Cloud Functions for feeds and notifications, plus QR-based sign-in using time-limited OAuth/JWT tokens.',
'Built location-based discovery and messaging with SwiftUI, MapKit, and Firebase Auth/Firestore/FCM. Implemented server-side feed aggregation and notifications in TypeScript Cloud Functions, plus QR-based cross-device sign-in with time-limited OAuth/JWT tokens.':'Built SwiftUI/MapKit discovery and messaging with Firebase. Implemented TypeScript Cloud Functions for feeds and notifications, plus QR-based sign-in using time-limited OAuth/JWT tokens.',
'Built an iOS app and Python/Flask backend for personalized glucose monitoring.':'Built an iOS application and Python/Flask backend for glucose monitoring.',
}
def subst(s):
 for a,b in replacements.items():s=s.replace(a,b)
 return s

def esc(s):
 return ''.join({'\\':r'\textbackslash{}','&':r'\&','%':r'\%','$':r'\$','#':r'\#','_':r'\_','{':r'\{','}':r'\}','~':r'\textasciitilde{}','^':r'\textasciicircum{}'}.get(c,c) for c in s)
def node(n):
 if n.tag==qn('w:r'):
  s=esc(subst(''.join(t.text or '' for t in n.iter(qn('w:t')))))
  return r'\textbf{'+s+'}' if n.find('.//'+qn('w:b')) is not None else s
 if n.tag==qn('w:hyperlink'):
  url=D.part.rels[n.get(qn('r:id'))].target_ref
  return r'\href{'+url+'}{'+''.join(node(c) for c in n)+'}'
 return ''
head=r'''% Compile with: tectonic Moein-Razavi-Unified-Resume.tex
% Business metrics are supplied and confirmed by the candidate.
% No forecasting metric, evaluation period, or people-management claim is inferred.
\documentclass[letterpaper,11pt]{article}
\usepackage[margin=0.7in,top=0.6in,bottom=0.6in]{geometry}
\usepackage{fontspec}
\setmainfont{Arial}
\usepackage[unicode,hidelinks]{hyperref}
\usepackage{enumitem}
\usepackage{needspace}
\pagestyle{empty}
\setlength{\parindent}{0pt}
\setlength{\parskip}{3.5pt}
\setlength{\emergencystretch}{2em}
\setlist[itemize]{leftmargin=11pt,label=\textbullet,itemsep=3.5pt,parsep=0pt,topsep=2pt,partopsep=0pt}
\newcommand{\ressection}[1]{\par\addvspace{8pt}\needspace{3\baselineskip}{\fontsize{11.5}{14}\selectfont\bfseries #1}\par\nobreak}
\newcommand{\employer}[1]{\par\addvspace{3pt}\needspace{4\baselineskip}\textbf{#1}\par\nobreak}
\hypersetup{pdftitle={Moein Razavi - AI/ML Engineer},pdfauthor={Moein Razavi}}
\begin{document}
\fontsize{10.5}{12.7}\selectfont
\raggedright
'''
lines=[head];inlist=False
for i,p in enumerate(D.paragraphs):
 text=''.join(t.text or '' for t in p._p.iter(qn('w:t')))
 style=p.style.name
 if not text:continue
 bullet=style=='List Bullet'
 if inlist and not bullet:lines.append(r'\end{itemize}');inlist=False
 if i==0:lines.append(r'{\fontsize{22}{26}\selectfont\bfseries '+esc(text)+r'}\par');continue
 if text=='Selected Projects':lines.append(r'\newpage')
 if style=='Heading 1':lines.append(r'\ressection{'+esc(text)+'}');continue
 if style=='Heading 2':lines.append(r'\employer{'+esc(text)+'}');continue
 if bullet and not inlist:lines.append(r'\begin{itemize}');inlist=True
 body=''.join(node(c) for c in p._p)
 if text.startswith('OpenSync:'):
  body=r'\textbf{OpenSync:} Co-developed open-source Python software using Lab Streaming Layer to synchronize sensors and experiment events; published tests measured RMS recording-time lag below 0.07 ms. \href{https://github.com/TAMUCogLab/OpenSync}{GitHub repository}.'
 if bullet:lines.append(r'\item '+body)
 else:lines.append(body+r'\par')
if inlist:lines.append(r'\end{itemize}')
lines.append(r'\end{document}')
Path('output/resume/Moein-Razavi-Unified-Resume.tex').write_text('\n'.join(lines))

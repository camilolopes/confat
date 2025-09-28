
import subprocess, re, sys, json, datetime as dt, os
from pathlib import Path

def run(cmd):
    return subprocess.check_output(cmd, text=True).strip()

def latest_tag():
    try:
        return run(["git", "describe", "--tags", "--abbrev=0"])
    except subprocess.CalledProcessError:
        return ""

def parse_semver(tag):
    m = re.match(r"v(\d+)\.(\d+)\.(\d+)$", tag or "")
    if not m:
        return (1, 0, 0)
    return tuple(map(int, m.groups()))

def collect_commits(since_tag):
    rng = f"{since_tag}..HEAD" if since_tag else "HEAD"
    try:
        raw = run(["git", "log", "--pretty=%H%x01%s%x01%an", rng])
        lines = [tuple(x.split("\x01")) for x in raw.splitlines() if x.strip()]
        commits = [{"hash":h, "subject":s, "author":a} for h,s,a in lines]
        return commits
    except subprocess.CalledProcessError:
        return []

def bump_by_commits(base_tag, commits):
    majors = any("#major" in c["subject"].lower() for c in commits)
    minors = any("#minor" in c["subject"].lower() for c in commits)
    M,m,p = parse_semver(base_tag)
    if not base_tag:
        return "v1.0.0"
    if majors:
        return f"v{M+1}.0.0"
    if minors:
        return f"v{M}.{m+1}.0"
    return f"v{M}.{m}.{p+1}"

def build_release_body(tag, commits):
    if not commits:
        return f"Automated release {tag}."
    lines = []
    for c in commits:
        subj = c["subject"].strip()
        # Strip bump flags from subject for cleaner notes
        subj = re.sub(r"\s+#(major|minor|patch)\b", "", subj, flags=re.I)
        lines.append(f"- {subj} ({c['author']})")
    return "### Changes\n" + "\n".join(lines)

def insert_versions_section(tag, body_text):
    path = Path("VERSIONS.md")
    if not path.exists():
        path.write_text("# 📘 VERSIONS.md — Faturas Processor App\n\n## 📈 Histórico de Versões\n", encoding="utf-8")
    md = path.read_text(encoding="utf-8")
    date_br = dt.date.today().strftime("%d/%m/%Y")
    section = f"\n### {tag} — {date_br}\n{body_text}\n\n---\n"
    anchor = "## 📈 Histórico de Versões"
    idx = md.find(anchor)
    if idx == -1:
        new_md = section + "\n" + md
    else:
        endline = md.find("\n", idx)
        if endline == -1: endline = len(md)
        new_md = md[:endline+1] + section + md[endline+1:]
    path.write_text(new_md, encoding="utf-8")

def main():
    base = latest_tag()
    commits = collect_commits(base)
    tag = bump_by_commits(base, commits)
    body = build_release_body(tag, commits)
    # Write release body for the next step
    with open("release_body.txt", "w", encoding="utf-8") as f:
        f.write(body)
    # Update VERSIONS.md immediately on main
    insert_versions_section(tag, body)
    # Export output to GITHUB_OUTPUT
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"tag={tag}\n")
    else:
        print(f"::set-output name=tag::{tag}")

if __name__ == "__main__":
    main()

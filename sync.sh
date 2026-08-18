#!/bin/bash
# Auto-sync local working copy -> github.com/forumdata-collab/gem-backup
# Run from cron; commits+pushes only when something changed.
cd /home/ubuntu/gembackup_web || exit 1
git add -A
if git diff --cached --quiet; then exit 0; fi
git -c user.name=ForumData -c user.email=forumdata@gmail.com commit -m "sync $(date -u +%Y-%m-%dT%H:%M:%SZ)" -q
git push origin main >/dev/null 2>&1 && echo "pushed $(git rev-parse --short HEAD)" || echo "PUSH FAILED"

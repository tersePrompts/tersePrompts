---
name: resume-docx-workflow
description: Working with resume documents in docx format, making updates via Word COM automation, and git best practices
---

# Resume DOCX Workflow

## Document Editing Strategies

### When XML/JSON libraries fail
- Use Word COM automation (PowerShell) for direct document manipulation
- Patterns that work:
  ```powershell
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open("path\to\file.docx")
  $doc.Content.Find.Execute("old text", $false, $false, $false, $false, $false, $true, 1, $false, "new text")
  $doc.Save()
  $doc.Close()
  $word.Quit()
  ```

### Key learnings
- docx XML is complex; text is split across `<w:t>` tags
- Simple string replacement often fails in XML
- Word COM automation works for find/replace operations
- en-dash (–) vs hyphen (-) matters in find patterns
- Selection-based edits: `$selection = $word.Selection`

## Resume Content Best Practices

### Business language vs technical
- Write projects in functional/business terms, not technical jargon
- Focus on: problem → solution → impact
- Use metrics: "80% faster", "60% cost reduction", "8 enterprise use cases"
- Make interviewers curious: "How did you solve that?"

### Typical resume sections
- Contact info
- Quote (optional)
- WORK EXPERIENCE with timeline
- Projects (business language)
- EDUCATION
- KEY SKILLS (technical + protocols)
- JOB SEARCH PREFERENCES

## Git Workflow

### Commit patterns
```bash
git status                          # Check what changed
git add file1.docx file2.pdf       # Stage files
git rm oldfile.pdf                 # Remove deleted files
git commit -m "descriptive message"
git push                           # Push to remote
```

### Commit message format
```
Update resume: Brief summary

- Change 1: specific detail
- Change 2: specific detail
- Impact: what this achieves
- Removed: old files
```

## Security Practices

### NEVER commit
- Passwords/tokens
- API keys
- Personal access tokens
- `.env` files
- Temporary Word files (`~$*.docx`)

### Store in memory instead
```markdown
---
name: github-token
description: GitHub Personal Access Token
metadata:
  type: user
---

Token: [actual token here]
Username: [username]
```

### Memory file location
`C:\Users\av201\.claude\projects\C--Users-av201-workspace-tersePrompts\memory\`

## Common Issues

### Token authentication fails (403)
- Token may lack `repo` permission
- User may not have repo access
- Try SSH instead: `git push`
- Generate new token with correct scopes

### Word COM automation errors
- "Bad parameter" → Wrong constant value
- COMException → Object doesn't exist or wrong method
- Text not found → Pattern doesn't match (check dash characters)

### File locked (EBUSY)
- File is open in Word - close it first
- Use different filename for output

## Workflow Summary

1. Copy original docx
2. Use Word COM for precise edits
3. Verify changes by extracting XML
4. Add to git, commit with descriptive message
5. Store credentials in memory, never in commits
6. Push with authenticated token

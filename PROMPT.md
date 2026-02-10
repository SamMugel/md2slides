@prd.md @activity.md

We are building the project according to the PRD in this repo.

First read activity.md to see what was recently accomplished.

## Start the Application

This is a Python CLI tool (no dev server needed). To test the CLI:
```bash
pip install -e .
python -m md2slides --help
```

To run a conversion:
```bash
python -m md2slides examples/simple.md -o output.pptx
```

## Work on Tasks

Open prd.md and find the single highest priority task where `"passes": false`.

Work on exactly ONE task:
1. Implement the change according to the task steps
2. Run any available checks:
   - `python -m py_compile md2slides/*.py` (syntax check)
   - `python -m md2slides --help` (verify CLI works)

## Verify Output

After implementing, verify your work:

1. Run the CLI with a test file:
   ```bash
   python -m md2slides examples/simple.md -o test-output.pptx
   ```

2. Verify the PPTX was created:
   ```bash
   ls -la test-output.pptx
   ```

3. For visual verification, you can open the PPTX in a viewer or use agent-browser if there's a preview available.

## Log Progress

Append a dated progress entry to activity.md describing:
- What you changed
- What commands you ran
- The screenshot filename
- Any issues encountered and how you resolved them

## Update Task Status

When the task is confirmed working, update that task's `"passes"` field in prd.md from `false` to `true`.

## Commit Changes

Make one git commit for that task only with a clear, descriptive message:
```
git add .
git commit -m "feat: [brief description of what was implemented]"
```

Do NOT run `git init`, do NOT change git remotes, and do NOT push.

## Important Rules

- ONLY work on a SINGLE task per iteration
- Always verify in browser before marking a task as passing
- Always log your progress in activity.md
- Always commit after completing a task

## Completion

When ALL tasks have `"passes": true`, output:

<promise>COMPLETE</promise>

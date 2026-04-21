Analyze staged or recent changes and write a concise conventional commit message.
Output ONLY the commit message text in chat. Do NOT run git commands.

Subject line rules:
- Format: <type>(<optional scope>): <description>
- Limit to 50 characters total.
- Capitalize the first letter of the description.
- Do not end with a period.
- Use imperative mood (e.g., "Add feature" not "Added feature").

Body rules:
- Separate from the subject with a single blank line.
- Bullet list of key changes.
- Wrap at 72 characters to ensure readability in CLI tools.

General rules:
- Output the commit message as plain text. Do NOT wrap it in backticks or code blocks.
- Keep commits atomic and focused -- each commit should represent a single, logical change.

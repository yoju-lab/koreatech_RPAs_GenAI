# Setting up a Daily Cron Job for News Analysis Script

This guide explains how to set up a cron job that runs the news analysis script (`03_analysis_with_news_openais.py`) once per day.

## Prerequisites

- Access to a Linux/Unix/macOS system with cron service
- Python environment properly set up for the script
- Appropriate permissions to edit crontab

## Steps to Set Up a Cron Job

### 1. Open the Crontab Editor

Open your terminal and type:

```bash
crontab -e
```

This command opens the crontab editor. If this is your first time using crontab, you may be prompted to select an editor (nano, vim, etc.).

### 2. Add the Cron Job Entry

Add the following line to schedule the script to run daily at a specific time (for example, 2:00 AM):

```bash
0 2 * * * /usr/bin/python3 /apps/koreatech_RPAs_GenAI/codes/06_analysis_openais/03_analysis_with_news_openais.py >> /apps/koreatech_RPAs_GenAI/logs/news_analysis_$(date +\%Y\%m\%d).log 2>&1
```

#### Explanation of the Cron Schedule Format:

- `0`: Minute (0)
- `2`: Hour (2 AM)
- `*`: Day of month (every day)
- `*`: Month (every month)
- `*`: Day of week (every day of the week)

#### Other Options:

- To run at 8:30 AM daily: `30 8 * * *`
- To run at midnight daily: `0 0 * * *`
- To run at 10:15 PM daily: `15 22 * * *`

The output and any errors will be redirected to a log file with the current date in the filename.

### 3. Save and Exit

Save the file and exit the editor:
- For nano: Press `Ctrl+O` to write the file, then `Ctrl+X` to exit
- For vim: Press `Esc`, then type `:wq` and press `Enter`

### 4. Verify the Cron Job

To verify that your cron job has been added:

```bash
crontab -l
```

This command lists all cron jobs for the current user.

## Troubleshooting

If the script doesn't run as expected:

1. Check if the Python path is correct. You may need to use the full path to the Python executable:
   ```bash
   which python3
   ```

2. Ensure the script has execution permissions:
   ```bash
   chmod +x /apps/koreatech_RPAs_GenAI/codes/06_analysis_openais/03_analysis_with_news_openais.py
   ```

3. Review the log file for any error messages.

4. Make sure any environment variables needed by the script are available in the cron environment. You might need to set them explicitly in the crontab or use a wrapper script.

## Important Notes

- Cron runs in a limited environment. If your script depends on environment variables, you may need to source them explicitly or use a wrapper script.
- Consider using absolute paths for all files and commands in your cron job.
- The `%` character is special in crontab - it needs to be escaped with a backslash (`\%`) if used in the command.

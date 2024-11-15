from rpa_bot.log import lte, log
from rpa_bot import Bot
import pandas as pd
from rpa_bot.bot_helpers import BotStatus


class Notifications():
    email = None
    password = None
    application_name = None
    base_url = "https://graph.microsoft.com/v1.0/me"

    def report_body(self, bot):
        import os
        import datetime
        try:
            subject = f"RPA Bot: {bot.bot_name} Finished"
            with open(os.path.join(bot.res_path, 'bot_mail.html'), "r") as template_file:
                text = template_file.read().replace('\n', '')
            text = text.replace('$robotVersion$', bot.config.bot_version)
            text = text.replace('$robotName$', bot.bot_name)
            now = datetime.datetime.now()
            now_formated = now.strftime("%Y.%m.%d %H:%M:%S")
            text = text.replace('$now$', now_formated)

            counters = bot.counters
            text = text.replace('$processed$', str(counters.processed))
            text = text.replace('$success$', str(counters.success))
            text = text.replace('$error$', str(counters.error))

            if bot.bot_status == BotStatus.success:
                robot_msg = "All cases were processed"
            elif bot.bot_status == BotStatus.warning:
                robot_msg = "Warning happend, please see the logs for details."
                subject = f"{subject} - warnings encountered"
            elif bot.bot_status == BotStatus.error:
                robot_msg = f"Error happend: [{bot.error}]"
                subject = f"{subject} - error encountered"
            else:
                robot_msg = "Unknown stop reason"
                subject = f"{subject} - unknown error appeared"

            text = text.replace('$stopReason$', robot_msg)
            return text
        except Exception as e:
            log(e, lte.error)
            raise e
    
    def exceptions_body(self, doc_list):
        try:
            text = "Dears,<br/><br/>For below document number(s) occured an error during document processing:<br/><br/>"
            df = pd.DataFrame.from_records(doc_list, columns=['Document number', 'Error'])
            html = df.to_html(index=False)
            text += html
            return text
        except Exception as e:
            log(e, lte.error)
            raise e
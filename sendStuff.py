# pandas, openpyxl

import pandas as pd
import win32com.client as win32
from pywintypes import com_error
import os
import sys

# location of the excel with the data
filePath = os.environ["FILEPATH"]

# columns to keep in the email attachment
COLUMNS = [
    "VIN",
    "Line",
    "Body Colour",
    "Roof Colour",
    "Interior Colour",
    "Type",
    "Agent",
    "Actual Location",
    "Arrival Zeebrugge",
    "Departure Zeebrugge",
    "Arrival Altishofen",
    "Status Galliker",
    "Arrival Agent",
    "Order Number",
    "Customer Name",
    "Payment Status",

]
# email subject
SUBJECT = "Planned Deliveries smart"

# don't touch stuff below

outlook = win32.Dispatch("outlook.application")


def sendMail(
    filename: str,
    # name: str = "guggus",
    target: dict,
    qc=True,
):
    """Prepare mail to agents, maybe send.

    Args:
        filename (str): Name of the file attachment.
        target (dict): Dictionary containing recipients.
        qc (bool, optional): Whether to just display (rather than try to send).
            Defaults to True.
    """
    # craete a new mail
    mail = outlook.CreateItem(0)

    # set recipient
    mail.To = target["to"]
    mail.CC = target["cc"]

    # set subject
    mail.Subject = SUBJECT

    # read from template file
    with open(os.path.dirname(__file__) + "\\" + "emailTemplate.txt", "r", encoding="utf-8") as f:
        body = f.read()

    # set the body of the message
    # replace name in template with input to this function
    # body = re.sub(r"\$name", name, body)

    mail.Body = body

    # attach file
    attachment = f"{agent}_orders_{pd.Timestamp.now().strftime('%d%b%Y')}.xlsx"

    mail.Attachments.Add(os.path.dirname(__file__) + "\\" + attachment)

    # show email in outlook
    mail.Display()
    if not qc:
        try:
            mail.Send()
        except com_error:
            print("Cannot send the mail because Outlook blocks me. Send it manually.")


if __name__ == "__main__":
    # read table
    dfs = pd.read_excel(filePath, sheet_name=["Assignment", "Agents"])
    df = dfs["Assignment"]
    df_agents = dfs["Agents"]

    # find location of actual information
    cellVIN = df.where(df == "VIN").dropna(how="all").dropna(axis=1)

    # figure out the headers
    headers = df.loc[cellVIN.index[0], cellVIN.columns[0] :].to_list()

    # keep only the relevant piece
    df = df.loc[cellVIN.index[0] + 1 :, cellVIN.columns[0] :]

    # reset the headers
    df.columns = headers
    # only send out the cars that have not yet been delivered
    df = df.loc[df["Type"] == "Customer",]
    # only send the columns that have been defined above
    df = df.loc[:, COLUMNS].reset_index(drop=True)

    for agent, df_loc in df.groupby("Agent"):
        # for every agent in the data
        if len(df_loc) > 0:
            # only do something if there was at least 1 order
            filename = f"{agent}_orders_{pd.Timestamp.now().strftime('%d%b%Y')}.xlsx"

            # save file with filter for this agent
            df_loc.to_excel(os.path.dirname(__file__) + "\\" + filename, index=False)

            # pre-allocate target dictionary
            target = dict()
            try:
                target["to"] = df_agents.loc[
                    df_agents["Agent"] == agent, "1st Contact"
                ].values[0]

                name1, name2 = df_agents.loc[df_agents["Agent"] == agent, "2nd Contact"].values[0], df_agents.loc[df_agents["Agent"] == agent, "3rd Contact"].values[0]
                ccs = f"{name1};{name2}"

                target["cc"] = ccs

            except IndexError:
                print(
                    f"!! Could not find a reference of '{agent}' in the Agents sheet !!"
                )
                sys.exit()

            sendMail(filename, target=target)
            #break
        else:
            print(f"{agent} did not have any new orders.")

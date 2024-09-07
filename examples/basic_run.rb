require 'mbox2csv'

# Define file paths
mbox_file = 'INBOX'
csv_file = 'emails.csv'
sender_stats_csv_file = 'email_statistics.csv'
recipient_stats_csv_file = 'recipient_statistics.csv'

# Initialize the parser with the file paths
parser = Mbox2CSV::MboxParser.new(mbox_file, csv_file, sender_stats_csv_file, recipient_stats_csv_file)

# Parse the MBOX file, save email data, and generate statistics
parser.parse

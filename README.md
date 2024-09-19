<h1 align="center">
    mbox2csv -- Parse MBOX files and export email data into CSV format
</h1>

<p align="center">
    <a href="#-installation">📦 Installation</a> •
    <a href="#-usage">🚀 Usage</a> •
    <a href="#-licence">🔑 Licence</a>
</p>

mbox2csv is a Ruby gem that provides a simple way to parse MBOX files and export email data into CSV format. It also generates valuable email statistics for data mining tasks, such as the number of emails sent by each sender and recipient and average body lengths. This is ideal for analyzing email datasets or processing email archives.

## 📦 Installation

```sh
$ gem install mbox2csv
```

## 🚀 Usage

### Basic run example

```ruby
require 'mbox2csv'

# Define file paths
mbox_file = '/path/to/the/INBOX_file'
all_emails = 'emails.csv'
sender_stats_all_emails = 'email_statistics.csv'
recipient_stats_all_emails = 'recipient_statistics.csv'

# Initialize the parser with the file paths
parser = Mbox2CSV::MboxParser.new(mbox_file, all_emails, sender_stats_all_emails, recipient_stats_all_emails)

# Parse the MBOX file, save email data, and generate statistics
parser.parse
```

## 🔑 License

This package is distributed under the MIT License. This license can be found online at <http://www.opensource.org/licenses/MIT>.

## Disclaimer

This framework is provided as-is, and there are no guarantees that it fits your purposes or that it is bug-free. Use it at your own risk!

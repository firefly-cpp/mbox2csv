require 'base64'
require 'csv'
require 'mail'
require 'fileutils'
require 'ruby-progressbar'

module Mbox2CSV
    # Main class for parsing MBOX files and saving email data and statistics to CSV files.
    class MboxParser
        # Initializes the MboxParser with file paths for the MBOX file, output CSV file,
        # and statistics CSV files for sender and recipient statistics.
        #
        # @param [String] mbox_file Path to the MBOX file to be parsed.
        # @param [String] csv_file Path to the output CSV file where parsed email data will be saved.
        # @param [String] stats_csv_file Path to the output CSV file where sender statistics will be saved.
        # @param [String] recipient_stats_csv_file Path to the output CSV file where recipient statistics will be saved.
        def initialize(mbox_file, csv_file, stats_csv_file, recipient_stats_csv_file)
            @mbox_file = mbox_file
            @csv_file = csv_file
            @statistics = EmailStatistics.new
            @stats_csv_file = stats_csv_file
            @recipient_stats_csv_file = recipient_stats_csv_file
            @senders_folder = 'senders/'
            FileUtils.mkdir_p(@senders_folder) # Create the senders folder if it doesn't exist
        end

        # Parses the MBOX file and writes the email data to the specified CSV file.
        # It also saves sender and recipient statistics to separate CSV files.
        # A progress bar is displayed during the parsing process.
        def parse
            total_lines = File.foreach(@mbox_file).inject(0) { |c, _line| c + 1 }
            progressbar = ProgressBar.create(title: "Parsing Emails", total: total_lines, format: "%t: |%B| %p%%")

            CSV.open(@csv_file, 'w') do |csv|
                csv << ['From', 'To', 'Subject', 'Date', 'Body']

                File.open(@mbox_file, 'r') do |mbox|
                    buffer = ""
                    mbox.each_line do |line|
                        progressbar.increment
                        if line.start_with?("From ")
                            process_email_block(buffer, csv) unless buffer.empty?
                            buffer = ""
                        end
                        buffer << line
                    end
                    process_email_block(buffer, csv) unless buffer.empty?
                end
            end
            puts "Parsing completed. Data saved to #{@csv_file}"

            @statistics.save_sender_statistics(@stats_csv_file)
            @statistics.save_recipient_statistics(@recipient_stats_csv_file)
        rescue => e
            puts "Error processing MBOX file: #{e.message}"
        end

        private

        # Processes an individual email block from the MBOX file, extracts the email fields,
        # and writes them to the CSV. Also records email statistics for analysis and creates
        # sender-specific CSV files.
        #
        # @param [String] buffer The email block from the MBOX file.
        # @param [CSV] csv The CSV object where email data is written.
        def process_email_block(buffer, csv)
            mail = Mail.read_from_string(buffer)

            from = ensure_utf8(mail.from ? mail.from.join(", ") : '', 'UTF-8')
            to = ensure_utf8(mail.to ? mail.to.join(", ") : '', 'UTF-8')
            subject = ensure_utf8(mail.subject ? mail.subject : '', 'UTF-8')
            date = ensure_utf8(mail.date ? mail.date.to_s : '', 'UTF-8')
            body = decode_body(mail)

            csv << [from, to, subject, date, body]

            @statistics.record_email(from, to, body.length)

            save_email_to_sender_csv(from, to, subject, date, body)
        rescue => e
            puts "Error processing email block: #{e.message}"
        end

        # Decodes the email body content based on content-transfer encoding and converts it to UTF-8.
        #
        # @param [Mail] mail The mail object to decode.
        # @return [String] The decoded email body.
        def decode_body(mail)
            body = if mail.multipart?
            part = mail.text_part || mail.html_part
            part&.body&.decoded || ''
        else
            mail.body.decoded
        end

        charset = mail.charset || 'UTF-8'

        case mail.content_transfer_encoding
        when 'base64'
            body = Base64.decode64(body)
        when 'quoted-printable'
            body = body.unpack('M').first
        end

        ensure_utf8(body, charset)
    end

    # Converts text to UTF-8 encoding, handling invalid characters by replacing them with '?'.
    #
    # @param [String] text The input text.
    # @param [String] charset The character set of the input text.
    # @return [String] UTF-8 encoded text.
    def ensure_utf8(text, charset)
        return '' if text.nil?
        text = text.force_encoding(charset) if charset
        text.encode('UTF-8', invalid: :replace, undef: :replace, replace: '?')
    end

    # Saves an email to a sender-specific CSV file.
    #
    # @param [String] from The sender of the email.
    # @param [String] to The recipient(s) of the email.
    # @param [String] subject The subject of the email.
    # @param [String] date The date of the email.
    # @param [String] body The body of the email.
    def save_email_to_sender_csv(from, to, subject, date, body)
        return if from.empty?

        sender_file = File.join(@senders_folder, "#{sanitize_filename(from)}.csv")

        CSV.open(sender_file, 'a') do |csv|
            if File.size(sender_file).zero?
                csv << ['From', 'To', 'Subject', 'Date', 'Body'] # Add header if file is new
            end
            csv << [from, to, subject, date, body]
        end
    rescue => e
        puts "Error writing to sender CSV file for #{from}: #{e.message}"
    end

    # Sanitizes filenames by replacing invalid characters with underscores.
    #
    # @param [String] filename The input filename.
    # @return [String] A sanitized version of the filename.
    def sanitize_filename(filename)
        filename.gsub(/[^0-9A-Za-z\-]/, '_')
    end
end

# The EmailStatistics class is responsible for gathering and writing statistics related to emails.
# It tracks sender frequency, recipient frequency, and calculates the average email body length per sender.
class EmailStatistics
    def initialize
        @sender_counts = Hash.new(0)
        @recipient_counts = Hash.new(0)
        @body_lengths = Hash.new { |hash, key| hash[key] = [] }
    end

    # Records an email's sender, recipients, and body length for statistical purposes.
    #
    # @param [String] from The sender of the email.
    # @param [String, Array<String>] to The recipient(s) of the email.
    # @param [Integer] body_length The length of the email body in characters.
    def record_email(from, to, body_length)
        return if from.empty?

        @sender_counts[from] += 1
        @body_lengths[from] << body_length

        Array(to).each do |recipient|
            @recipient_counts[recipient] += 1
        end
    end

    # Saves sender statistics to a CSV file and prints them to the console.
    #
    # @param [String] csv_filename The path to the output CSV file for sender statistics.
    def save_sender_statistics(csv_filename)
        sorted_senders = @sender_counts.sort_by { |_sender, count| -count }
        average_body_lengths = @body_lengths.transform_values { |lengths| lengths.sum / lengths.size.to_f }

        CSV.open(csv_filename, 'w') do |csv|
            csv << ['Sender', 'Email Count', 'Average Body Length (chars)']
            sorted_senders.each do |sender, count|
                avg_length = average_body_lengths[sender].round(2)
                csv << [sender, count, avg_length]
            end
        end

        puts "Sender Email Statistics:"
        sorted_senders.each do |sender, count|
            avg_length = average_body_lengths[sender].round(2)
            puts "#{sender}: #{count} emails, Average body length: #{avg_length} chars"
        end
    end

    # Saves recipient statistics to a CSV file and prints them to the console.
    #
    # @param [String] csv_filename The path to the output CSV file for recipient statistics.
    def save_recipient_statistics(csv_filename)
        sorted_recipients = @recipient_counts.sort_by { |_recipient, count| -count }

        CSV.open(csv_filename, 'w') do |csv|
            csv << ['Recipient', 'Email Count']
            sorted_recipients.each do |recipient, count|
                csv << [recipient, count]
            end
        end

        puts "\nRecipient Email Statistics:"
        sorted_recipients.each do |recipient, count|
            puts "#{recipient}: #{count} emails"
        end
    end
end
end

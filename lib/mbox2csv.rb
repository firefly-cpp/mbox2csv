require 'base64'
require 'csv'
require 'mail'
require 'fileutils'
require 'ruby-progressbar'

module Mbox2CSV
    # Main class for parsing MBOX files, saving email data/statistics to CSV,
    # and (optionally) extracting selected attachment types to disk.
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

        # Extract selected attachment file types from the MBOX into a folder.
        #
        # @param [Boolean] extract         Flag to enable/disable extraction.
        # @param [Array<String>] filetypes Array of extensions to extract (e.g., %w[pdf jpg png]).
        # @param [String] output_folder    Directory to write attachments into.
        # @return [Integer]                Number of files successfully written.
        def extract_attachments(extract: true, filetypes: [], output_folder: 'attachments')
            return 0 unless extract

            wanted_exts = Array(filetypes).map { |e| e.to_s.downcase.sub(/\A\./, '') }.uniq
            raise ArgumentError, "filetypes must not be empty when extract: true" if wanted_exts.empty?

            FileUtils.mkdir_p(output_folder)
            total_written = 0

            total_lines = File.foreach(@mbox_file).inject(0) { |c, _| c + 1 }
            progressbar = ProgressBar.create(title: "Extracting Attachments", total: total_lines, format: "%t: |%B| %p%%")

            File.open(@mbox_file, 'r') do |mbox|
                buffer = ""
                mbox.each_line do |line|
                    progressbar.increment
                    if line.start_with?("From ")
                        total_written += process_attachment_block(buffer, wanted_exts, output_folder) unless buffer.empty?
                        buffer = ""
                    end
                    buffer << line
                end
                total_written += process_attachment_block(buffer, wanted_exts, output_folder) unless buffer.empty?
            end

            puts "Attachment extraction completed. #{total_written} file(s) saved to #{output_folder}"
            total_written
        rescue => e
            puts "Error extracting attachments: #{e.message}"
            0
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
            body =
                    if mail.multipart?
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
            if File.size?(sender_file).nil? || File.size(sender_file).zero?
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

    # --- Helpers for attachment extraction ---

    # Process a single email block to extract wanted attachments.
    def process_attachment_block(buffer, wanted_exts, output_folder)
        return 0 if buffer.nil? || buffer.empty?

        mail = Mail.read_from_string(buffer)
        return 0 unless mail

        written = 0
        date = (mail.date rescue nil)
        date_str = date ? date.strftime("%Y-%m-%d") : "unknown_date"
        time_str = date ? date.strftime("%H-%M-%S") : "unknown_time"

        Array(mail.attachments).each do |att|
            begin
                original_name = att.filename || att.name || "attachment"
                base = File.basename(original_name, ".*")
                ext  = File.extname(original_name).downcase.sub(/\A\./, '')

                # If no ext present, try to infer from MIME type
                ext = mime_to_ext(att.mime_type) if ext.empty? && att.mime_type

                # Skip if extension not desired
                next unless wanted_exts.include?(ext.downcase)

                safe_base = sanitize_filename(base)
                fname = "#{safe_base}_#{date_str}_#{time_str}.#{ext}"
                path  = File.join(output_folder, fname)

                # Ensure uniqueness if file already exists
                path = uniquify_path(path)

                # Write decoded content
                File.open(path, "wb") { |f| f.write(att.body.decoded) }
                written += 1
            rescue => e
                puts "Failed to save attachment '#{att&.filename}': #{e.message}"
            end
        end

        written
    rescue => e
        puts "Error processing attachment block: #{e.message}"
        0
    end

    # Minimal MIME→extension mapping; extend as needed.
    def mime_to_ext(mime)
        map = {
            'application/pdf' => 'pdf',
            'image/jpeg'      => 'jpg',
            'image/jpg'       => 'jpg',
            'image/png'       => 'png',
            'image/gif'       => 'gif',
            'text/plain'      => 'txt',
            'application/zip' => 'zip',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document' => 'docx',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' => 'xlsx',
            'application/msword' => 'doc',
            'application/vnd.ms-excel' => 'xls'
        }
        map[mime] || 'bin'
    end

    # If a path exists, append an incrementing suffix before the extension.
    def uniquify_path(path)
        return path unless File.exist?(path)
        dir  = File.dirname(path)
        base = File.basename(path, ".*")
        ext  = File.extname(path)
        i = 1
        new_path = File.join(dir, "#{base}_#{i}#{ext}")
        while File.exist?(new_path)
            i += 1
            new_path = File.join(dir, "#{base}_#{i}#{ext}")
        end
        new_path
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

# --- Usage example ---
# parser = Mbox2CSV::MboxParser.new("inbox.mbox", "emails.csv", "sender_stats.csv", "recipient_stats.csv")
# parser.parse
# parser.extract_attachments(extract: true, filetypes: %w[pdf jpg], output_folder: "exports")

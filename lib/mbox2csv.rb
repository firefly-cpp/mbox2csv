require 'base64'
require 'csv'
require 'mail'

module Mbox2CSV
    # Main class
    class MboxParser
        def initialize(mbox_file, csv_file, stats_csv_file, recipient_stats_csv_file)
            @mbox_file = mbox_file
            @csv_file = csv_file
            @statistics = EmailStatistics.new
            @stats_csv_file = stats_csv_file
            @recipient_stats_csv_file = recipient_stats_csv_file
        end

        def parse
            CSV.open(@csv_file, 'w') do |csv|
                # Write CSV header
                csv << ['From', 'To', 'Subject', 'Date', 'Body']

                File.open(@mbox_file, 'r') do |mbox|
                    buffer = ""
                    mbox.each_line do |line|
                        if line.start_with?("From ")
                            process_email_block(buffer, csv) unless buffer.empty?
                            buffer = "" # Reset buffer
                        end
                        buffer << line # Append line to buffer
                    end
                    process_email_block(buffer, csv) unless buffer.empty? # Process last email block
                end
            end
            puts "Parsing completed. Data saved to #{@csv_file}"

            # Save and print statistics after parsing
            @statistics.save_sender_statistics(@stats_csv_file)
            @statistics.save_recipient_statistics(@recipient_stats_csv_file)
        rescue => e
            puts "Error processing MBOX file: #{e.message}"
        end

        private

        def process_email_block(buffer, csv)
            mail = Mail.read_from_string(buffer)

            from = ensure_utf8(mail.from ? mail.from.join(", ") : '', 'UTF-8')
            to = ensure_utf8(mail.to ? mail.to.join(", ") : '', 'UTF-8')
            subject = ensure_utf8(mail.subject ? mail.subject : '', 'UTF-8')
            date = ensure_utf8(mail.date ? mail.date.to_s : '', 'UTF-8')

            body = decode_body(mail)

            # Write to CSVv
            csv << [from, to, subject, date, body]

            @statistics.record_email(from, to, body.length)
        rescue => e
            puts "Error processing email block: #{e.message}"
        end

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

    def ensure_utf8(text, charset)
        return '' if text.nil?
        text = text.force_encoding(charset) if charset
        text.encode('UTF-8', invalid: :replace, undef: :replace, replace: '?')
    end
end

class EmailStatistics
    def initialize
        @sender_counts = Hash.new(0) # Keeps count of emails per sender
        @recipient_counts = Hash.new(0) # Keeps count of emails per recipient
        @body_lengths = Hash.new { |hash, key| hash[key] = [] } # Stores body lengths per sender
    end

    def record_email(from, to, body_length)
        return if from.empty?

        @sender_counts[from] += 1
        @body_lengths[from] << body_length

        Array(to).each do |recipient|
            @recipient_counts[recipient] += 1
        end
    end

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

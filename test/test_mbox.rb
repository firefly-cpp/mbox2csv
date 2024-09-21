require_relative '../lib/mbox2csv'
require 'minitest/autorun'
require 'csv'

class MboxParserTest < Minitest::Test
    def setup
        # test files
        @mbox_file = 'test/test_simple.mbox'
        @all_emails = 'test/emails.csv'
        @sender_stats_all_emails = 'test/sender_stats.csv'
        @recipient_stats_all_emails = 'test/recipient_stats.csv'

        # Init parser for the test file
        @parser = Mbox2CSV::MboxParser.new(@mbox_file, @all_emails, @sender_stats_all_emails, @recipient_stats_all_emails)

        # Parse the MBOX file and generate statistics for testing
        @parser.parse
    end

    def test_email_parsing
        assert File.exist?(@all_emails), "CSV file should exist after parsing."

        # row count
        email_data = CSV.read(@all_emails, headers: true)
        assert_equal 5, email_data.size, "There should be 5 parsed emails."

        # headers
        expected_headers = ["From", "To", "Subject", "Date", "Body"]
        assert_equal expected_headers, email_data.headers, "CSV headers should match the expected email headers."
    end

    # Test case to check total senders
    def test_total_senders
        assert File.exist?(@sender_stats_all_emails), "Sender statistics CSV should exist after parsing."

        sender_data = CSV.read(@sender_stats_all_emails, headers: true)
        assert_equal 3, sender_data.size, "There should be 3 unique senders in the sender statistics file."

        # Check if a specific sender's email count is correct
        sender_info = sender_data.find { |row| row["Sender"] == "test@example.com" }
        assert_equal "2", sender_info["Email Count"], "Sender 'test@example.com' should have sent 2 emails."
    end

    # Test case to check total recipients
    def test_total_recipients
        assert File.exist?(@recipient_stats_all_emails), "Recipient statistics CSV should exist after parsing."

        recipient_data = CSV.read(@recipient_stats_all_emails, headers: true)
        assert_equal 4, recipient_data.size, "There should be 4 unique recipients in the recipient statistics file."

        # Check if a specific recipient's email count is correct (mock test data)
        recipient_info = recipient_data.find { |row| row["Recipient"] == "recipient@example.com" }
        assert_equal "2", recipient_info["Email Count"], "Recipient 'recipient@example.com' should have received 2 emails."
    end

end


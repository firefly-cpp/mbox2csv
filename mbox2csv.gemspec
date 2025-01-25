# frozen_string_literal: true

Gem::Specification.new do |spec|
    spec.name          = 'mbox2csv'
    spec.version       = '0.1.2'
    spec.license       = 'MIT'
    spec.authors       = %w[firefly-cpp]
    spec.email         = ['iztok@iztok-jr-fister.eu']

    spec.summary       = 'Parse MBOX files and export email data into CSV format'
    spec.homepage      = 'https://github.com/firefly-cpp/mbox2csv'
    spec.required_ruby_version = '>= 2.6.0'

    spec.metadata['homepage_uri'] = spec.homepage
    spec.metadata['source_code_uri'] = 'https://github.com/firefly-cpp/mbox2csv'
    spec.metadata['changelog_uri'] = 'https://github.com/firefly-cpp/mbox2csv'

    spec.files         = Dir["lib/**/*.rb"] + ["README.md", "LICENSE"]
    spec.require_paths = ['lib']

    spec.add_dependency "base64", "~> 0.2.0"
    spec.add_dependency "csv", "~> 3.3"
    spec.add_dependency "mail", "~> 2.8.1"
    spec.add_dependency "ruby-progressbar", "~> 1.11"

end

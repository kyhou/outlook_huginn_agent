# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

Gem::Specification.new do |spec|
  spec.name          = "huginn_outlook_agent"
  spec.version       = '0.1'
  spec.authors       = ["kyhou"]
  spec.email         = ["joelskill@gmail.com"]

  spec.summary       = %q{Huginn agent for Microsoft Outlook email integration}
  spec.description   = %q{Enables Huginn to interact with Microsoft Outlook email via Microsoft Graph API for receiving and sending emails}

  spec.homepage      = "https://github.com/[my-github-username]/huginn_outlook_agent"

  spec.license       = "MIT"


  spec.files         = Dir['LICENSE.txt', 'lib/**/*']
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = Dir['spec/**/*.rb'].reject { |f| f[%r{^spec/huginn}] }
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 2.0"
  spec.add_development_dependency "rake", "~> 13.0"

  spec.add_runtime_dependency "huginn_agent"
  spec.add_runtime_dependency "httparty", "~> 0.20"
  spec.add_runtime_dependency "json", "~> 2.6"
  spec.add_runtime_dependency "oauth2", "~> 2.0"
end

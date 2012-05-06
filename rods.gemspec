# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = %q{rods}
  s.version = "0.9.1"

  s.required_rubygems_version = Gem::Requirement.new(">= 1.2") if s.respond_to? :required_rubygems_version=
  s.authors = ["Dr. Heinz Breinlinger"]
  s.date = %q{2011-01-17}
  s.description = %q{OpenOffice.org oocalc: Fast automated batch-processing of spreadsheets (*.ods) conforming to Open Document Format v1.1. used by e.g. OpenOffice.org and LibreOffice. Please see screenshot and Rdoc-Documentation at http://ruby.homelinux.com/ruby/rods/. You can contact me at rods.ruby@online.de (and tell me about your experiences or drop me a line, if you like it ;-)}
  s.email = %q{rods.ruby@online.de}
  s.extra_rdoc_files = ["README", "lib/rods.rb"]
  s.files = ["Manifest", "README", "Rakefile", "lib/rods.rb", "rods.gemspec"]
  s.homepage = %q{http://ruby.homelinux.com/ruby/rods/}
  s.rdoc_options = ["--line-numbers", "--inline-source", "--title", "Rods", "--main", "README"]
  s.require_paths = ["lib"]
  s.rubyforge_project = %q{rods}
  s.rubygems_version = %q{1.3.7}
  s.summary = %q{Automation of OpenOffice/LibreOffice by batch-processing of spreadsheets conforming to Open Document v1.1}

  if s.respond_to? :specification_version then
    current_version = Gem::Specification::CURRENT_SPECIFICATION_VERSION
    s.specification_version = 3

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
    else
    end
  else
  end
end

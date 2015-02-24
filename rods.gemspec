# -*- encoding: utf-8 -*-
Gem::Specification.new do |s|
  s.name = %q{rods}
  s.version = "0.9.1"
  s.platform = Gem::Platform::RUBY
  s.has_rdoc = true
  s.extra_rdoc_files = %W(README) + Dir["lib/**/*.rb"]
  s.rdoc_options = ["--line-numbers", "--inline-source", "--title", "Rods", "--main", "README"]
  s.summary = %q{Automation of OpenOffice/LibreOffice by batch-processing of spreadsheets conforming to Open Document v1.1}
  s.description = %q{OpenOffice.org oocalc: Fast automated batch-processing of spreadsheets (*.ods) conforming to Open Document Format v1.1. used by e.g. OpenOffice.org and LibreOffice. Please see screenshot and Rdoc-Documentation at http://ruby.homelinux.com/ruby/rods/. You can contact me at rods.ruby@online.de (and tell me about your experiences or drop me a line, if you like it ;-)}
  s.author = "Dr. Heinz Breinlinger"
  s.email = %q{rods.ruby@online.de}
  s.homepage = %q{http://ruby.homelinux.com/ruby/rods/}
  s.files = %W(README Rakefile rods.gemspec) + Dir["{spec,lib}/**/*.rb"]
  s.require_path = "lib"
  s.add_runtime_dependency = "zip"
end

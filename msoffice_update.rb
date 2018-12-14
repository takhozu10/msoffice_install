#!/usr/bin/ruby
#
# # #
# MIT License
#
# Copyrigth (c) 2018 Tak Hozumi.
#
#   Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
#   associated documentation files (the "Software"), to deal in the Software without restriction, including
#   without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#   copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the
#   following conditions:
#
#   The above copyright notice and this permission notice shall be included in all copies or substantial
#   portions of the Software.
#
#   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
#   NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
#   IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
#   WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
#   SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
# # #

# # #
# Microsoft Office for Mac update script
#
#
#
# Written by: Tak Hozumi : End User Support Engineer : Big Lots Stores, Inc.
#
# Created on: Dec 13, 2018
# Modified on:
# Version: 1.0.0
# # #

# functions
def check_process(app_name)
  system("ps aux | grep '#{app_name}.app/Contents/MacOS/#{app_name}' | grep -v grep -c")
end

# variables
ms_word="Microsoft Word"
ms_excel="Microsoft Excel"
ms_pp="Microsoft PowerPoint"
ms_outlook="Microsoft Outlook"
ms_onenote="Microsoft OneNote"

# Run checks
if check_process(ms_word)
  puts "#{ms_word} is running on this computer"
else
  puts "#{ms_word} is not running"
end

if check_process(ms_excel)
  puts "#{ms_excel} is running on this computer"
else
  puts "#{ms_excel} is not running"
end

if check_process(ms_pp)
  puts "#{ms_pp} is running on this computer"
else
  puts "#{ms_pp} is not running"
end

if check_process(ms_outlook)
  puts "#{ms_outlook} is running on this computer"
else
  puts "#{ms_outlook} is not running"
end

if check_process(ms_onenote)
  puts "#{ms_onenote} is running on this computer"
else
  puts "#{ms_onenote} is not running"
end

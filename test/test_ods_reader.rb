# -*- coding: utf-8 -*-

require 'helper'

require 'lib/ods_reader'

class TestOdsReader < Test::Unit::TestCase

  def test_colreps
    rows = []
    ODSReader.open('test/files/colreps.ods', 'Sheet1') do |row|
      rows << row
    end
    assert_equal [["abc", "いろは", "123", "", "", "", ""],
      ["2009-12-06", "", "", "3", "5", "1", ""],
      ["98", "5", "5", "5", "1", "1", "2"]], rows
  end
end

require "busted.runner" ()
kpse.set_program_name "luatex"

local log = require "spreadsheet.log"
local xlsx = require "spreadsheet.xlsx-reader"

log.level="warn"
describe("Basic xlsx file loading should work", function()
  local lo,msg  = xlsx.load("test/pokus.xlsx")
  it("Parse the xlst file", function()
    assert.truthy(lo)
  end)
  local sheet = lo:get_sheet("Sheet1")
  it("Load a sheet", function()
    assert.truthy(sheet)
    assert.same(type(sheet), "table")
  end)
end)

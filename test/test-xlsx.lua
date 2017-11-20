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
  it("Should load data", function()
    assert.truthy(sheet.columns > 0)
    assert.same(#sheet.table, sheet.rows)
  end)
  local first = sheet.table[1]
  it("There should be correct number of columns", function()
    assert.same(sheet.columns, #first)
  end)
  it("Cells should have a correct form", function()
    local cell = first[1]
    assert.same(#cell, 1)
    assert.same(cell[1].value,"asdaf")
  end)

  local last = sheet.table[6]
  it("Complex shells should have the correct form as well", function()
    local cell = last[2]
    local t = {}
    assert.same(#cell, 3)
    for _, cell in ipairs(cell) do
      t[#t+1] =cell.value 
    end
    assert.same(table.concat(t), "jo a taky sloučíme buňky")
    assert.same(cell[2].style.color , "FFFF3333")

  end)
  it("Merged cells should be empty", function()
    assert.same(#last[3], 0)
  end)
  it("The empty row should be really empty", function()
    local empty = sheet.table[8]
    assert.same(#empty, 3)
    local s = ""
    for _, cell in ipairs(empty) do
      for _, x in ipairs(cell) do
        s = s .. x.value
      end
    end
    assert.same(s, "")
  end)
  it("Links should work", function()
    local cell = sheet.table[9][1]
    for _, x in ipairs(cell) do
      print(x.style.link, x.value)
    end
  end)
end)

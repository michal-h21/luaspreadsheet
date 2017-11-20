require "busted.runner"
kpse.set_program_name "luatex"

local log = require "spreadsheet.log"
local xlsx = require "spreadsheet.xlsx-reader"

log.level="warn"

local lo,msg  = xlsx.load("pokus.xlsx")

-- local obj,msg  = xslt.load("odpis.xlsx")

if not obj and not lo then
  print(msg)
  os.exit()
end

local sheet = lo:get_sheet("Sheet1")

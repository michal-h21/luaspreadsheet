-- module for range handling
local m = {}

local function get_range(range)
  local range = range:lower()
  local function getNumber(s)
    if s == "" or s == nil then return nil end
    local f,ex = 0,0
    for i in string.gmatch(s:reverse(),"(.)") do
      f = f + (i:byte()-96) * 26 ^ ex
      ex = ex + 1
    end
    return f
  end
  local x1,y1,x2,y2 =  range:match("(%a*)(%d*):*(%a*)(%d*)")
  return getNumber(x1),tonumber(y1),getNumber(x2),tonumber(y2)
end

m.get_range = get_range
return m

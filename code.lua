local function parse_number(v)
  if v == nil then return nil end
  local s = tostring(v):gsub("%s+", ""):gsub(",", ".")
  if s == "" then return nil end
  return tonumber(s)
end

local function parse_datetime(v)
  if v == nil then return nil end
  local s = tostring(v):gsub("^%s+", ""):gsub("%s+$", "")
  local d, m, y, hh, mm, ss =
    s:match("(%d%d)/(%d%d)/(%d%d%d%d)%s+(%d%d):(%d%d):(%d%d)")
  if d then
    return string.format("%s-%s-%sT%s:%s:%sZ", y, m, d, hh, mm, ss)
  end
  return nil
end

function run(ctx)
  ctx.log("XTB CASH OPERATION HISTORY importer start")

  local sheet = ctx.api.parse_xlsx("CASH OPERATION HISTORY")
  if not sheet or not sheet.rows then
    return { transactions = {} }
  end

  local rows = sheet.rows
  local header_row = 11

  local header = rows[header_row]
  if not header
    or header[2] ~= "ID"
    or header[3] ~= "Type"
    or header[4] ~= "Time"
    or header[5] ~= "Comment"
    or header[6] ~= "Symbol"
    or header[7] ~= "Amount"
  then
    ctx.log("Header mismatch at B11")
    return { transactions = {} }
  end

  local transactions = {}

  for i = header_row + 1, #rows do
    local r = rows[i]

    local typ = r[3]
    local time = parse_datetime(r[4])
    local comment = r[5]
    local symbol = r[6]
    local amount = parse_number(r[7])

    if not typ or not time or not comment or not symbol or not amount then
      goto continue
    end

    if typ == "Stock purchase" then
      local units, price =
        comment:match("OPEN BUY ([%d%.]+)/?[%d%.]* @ ([%d%.]+)")
      units = tonumber(units)
      price = tonumber(price)

      if units and price then
        local price_portfolio = math.abs(amount) / units
        table.insert(transactions, {
          ticker = "",
          trade_datetime = time,
          side = "buy",
          units = units,

          instrument_currency = nil,
          price_instrument = price,
          fx_rate = price_portfolio / price,
          price_portfolio = price_portfolio,
          total_portfolio = math.abs(amount),
          fee_portfolio = 0,
          tax_portfolio = 0,

          note = "XTB cash operation open",

          import_name = symbol
        })
      end

    elseif typ == "Stock sale" then
      local units, price =
        comment:match("CLOSE BUY ([%d%.]+)/?[%d%.]* @ ([%d%.]+)")
      units = tonumber(units)
      price = tonumber(price)

      if units and price then
        local price_portfolio = math.abs(amount) / units
        table.insert(transactions, {
          ticker = "",
          trade_datetime = time,
          side = "sell",
          units = units,

          instrument_currency = nil,
          price_instrument = price,
          fx_rate = price_portfolio / price,
          price_portfolio = price_portfolio,
          total_portfolio = math.abs(amount),
          fee_portfolio = 0,
          tax_portfolio = 0,

          note = "XTB cash operation close",

          import_name = symbol
        })
      end
    end

    ::continue::
  end

  ctx.log("Transactions created: " .. tostring(#transactions))
  return { transactions = transactions }
end

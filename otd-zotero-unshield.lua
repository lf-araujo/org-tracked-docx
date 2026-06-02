-- otd-zotero-unshield.lua
-- Pandoc Lua filter: replace ❰ZOTNNNN❱ placeholder tokens with RawInline OOXML
-- field-character runs sourced from a sidecar JSON written by otd_zotero_shield.py.
--
-- Usage (markdown -> docx):
--   pandoc input.md -t docx \
--     --metadata zotrefs-sidecar=/abs/path/to/source.zotrefs.json \
--     --lua-filter=/path/to/otd-zotero-unshield.lua \
--     -o output.docx
--
-- The sidecar must be the JSON produced by `otd_zotero_shield.py shield`,
-- containing {"source_docx": "...", "fields": {"NNNN": "<w:r>…</w:r>", ...}}.

local saved = {}
local stats = { restored = 0, missing = 0 }

local function load_sidecar(meta)
  local s = meta["zotrefs-sidecar"]
  if not s then
    io.stderr:write(
      "[zotero-unshield] no zotrefs-sidecar metadata supplied; placeholders will pass through as text\n")
    return
  end
  local path = pandoc.utils.stringify(s)
  local fh, err = io.open(path, "r")
  if not fh then
    io.stderr:write(string.format("[zotero-unshield] cannot open %q: %s\n", path, err or "?"))
    return
  end
  local raw = fh:read("*a")
  fh:close()
  local ok, data = pcall(pandoc.json.decode, raw)
  if not ok then
    io.stderr:write(string.format("[zotero-unshield] sidecar JSON parse failed: %s\n", tostring(data)))
    return
  end
  saved = data.fields or data
  local n = 0
  for _ in pairs(saved) do n = n + 1 end
  io.stderr:write(string.format("[zotero-unshield] loaded %d field chunks from %s\n", n, path))
end

local function unshield_str(elem)
  local text = elem.text
  if not text:find("❰ZOT", 1, true) then
    return nil
  end
  local out = {}
  local last = 1
  for s, key, e in text:gmatch("()❰ZOT(%d%d%d%d)❱()") do
    if s > last then
      table.insert(out, pandoc.Str(text:sub(last, s - 1)))
    end
    local chunk = saved[key]
    if chunk then
      table.insert(out, pandoc.RawInline("openxml", chunk))
      stats.restored = stats.restored + 1
    else
      table.insert(out, pandoc.Str(text:sub(s, e - 1)))
      stats.missing = stats.missing + 1
    end
    last = e
  end
  if last <= #text then
    table.insert(out, pandoc.Str(text:sub(last)))
  end
  return out
end

local function final_report(doc)
  io.stderr:write(string.format(
    "[zotero-unshield] restored=%d missing=%d\n",
    stats.restored, stats.missing))
  return doc
end

-- Two-pass filter: load metadata first, then transform Str inlines, then report.
return {
  { Meta = load_sidecar },
  { Str = unshield_str },
  { Pandoc = final_report },
}

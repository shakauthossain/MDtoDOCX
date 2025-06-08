function Table(el)
  el.attributes = el.attributes or {}
  el.attributes['style'] = 'width:100%;border:1px solid black;border-collapse:collapse'
  return el
end

function Cell(el)
  el.attributes = el.attributes or {}
  el.attributes['style'] = 'border:1px solid black;padding:6px;'
  return el
end

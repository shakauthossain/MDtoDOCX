function Table(elem)
  elem.attr = elem.attr or {}
  elem.attr.classes = elem.attr.classes or {}
  elem.attr.attributes = elem.attr.attributes or {}
  elem.attr.attributes['border'] = '1'
  return elem
end

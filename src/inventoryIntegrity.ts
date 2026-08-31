type InventoryTagRecord = { id: string; assetTag: string; categoryId: string }
type NamedInventoryTagRecord = InventoryTagRecord & { name: string }
type CategoryRecord = { id: string; name: string; prefix?: string }
type MigratableInventoryData = {
  categories: CategoryRecord[]
  items: NamedInventoryTagRecord[]
}

export const ASSET_TAG_PATTERN = /^[A-Z]{3}-\d{3}$/
export const BUILT_IN_CATEGORY_PREFIXES: Record<string, string> = {
  'cat-furniture': 'MOB',
  'cat-lighting': 'BEL',
  'cat-sound': 'LJU',
  'cat-kitchen': 'KOK',
}

export const normalizeAssetTag = (value: string) => value.trim().toUpperCase()

export const normalizeCategoryPrefix = (value: string) => value.trim().toUpperCase()

const nextCategoryPrefix = (name: string, usedPrefixes: Set<string>) => {
  const letters = name.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z]/g, '')
  const padded = `${letters}XXX`
  for (let second = 1; second < padded.length - 1; second += 1) {
    for (let third = second + 1; third < padded.length; third += 1) {
      const candidate = `${padded[0]}${padded[second]}${padded[third]}`
      if (!usedPrefixes.has(candidate)) return candidate
    }
  }
  for (let index = 0; index < 26; index += 1) {
    const candidate = `${padded.slice(0, 2)}${String.fromCharCode(65 + index)}`
    if (!usedPrefixes.has(candidate)) return candidate
  }
  return 'NYA'
}

export const migrateSampleInventory = <T extends MigratableInventoryData>(data: T): T => {
  const usedPrefixes = new Set<string>()
  const categories = data.categories.map((category) => {
    const configured = normalizeCategoryPrefix(BUILT_IN_CATEGORY_PREFIXES[category.id] ?? category.prefix ?? '')
    const prefix = /^[A-Z]{3}$/.test(configured) && !usedPrefixes.has(configured)
      ? configured
      : nextCategoryPrefix(category.name, usedPrefixes)
    usedPrefixes.add(prefix)
    return { ...category, prefix }
  })
  const prefixByCategory = new Map(categories.map((category) => [category.id, category.prefix]))
  const usedTags = new Set<string>()
  const items = data.items.map((item) => {
    const prefix = prefixByCategory.get(item.categoryId) ?? 'OVR'
    const normalizedTag = normalizeAssetTag(item.assetTag)
    const suffix = normalizedTag.match(/(\d{1,3})$/)?.[1]
    let number = suffix ? Number(suffix) : 1
    let assetTag = `${prefix}-${String(number).padStart(3, '0')}`
    for (let attempts = 0; usedTags.has(assetTag) && attempts < 999; attempts += 1) {
      number = number >= 999 ? 1 : number + 1
      assetTag = `${prefix}-${String(number).padStart(3, '0')}`
    }
    usedTags.add(assetTag)
    return { ...item, assetTag }
  })
  return { ...data, categories, items }
}

export const categoryPrefixFor = (categoryId: string, categories: CategoryRecord[]) => normalizeCategoryPrefix(BUILT_IN_CATEGORY_PREFIXES[categoryId] ?? categories.find((category) => category.id === categoryId)?.prefix ?? '')

export const assetTagError = (value: string, items: InventoryTagRecord[], categories: CategoryRecord[], currentItemId?: string, categoryId?: string) => {
  const assetTag = normalizeAssetTag(value)
  if (!ASSET_TAG_PATTERN.test(assetTag)) return 'Inventarienumret ska bestå av tre bokstäver, bindestreck och tre siffror, till exempel MOB-014.'
  if (items.some((item) => item.id !== currentItemId && normalizeAssetTag(item.assetTag) === assetTag)) return `Inventarienumret ${assetTag} används redan.`
  if (categoryId) {
    const prefix = assetTag.slice(0, 3)
    const expectedPrefix = categoryPrefixFor(categoryId, categories)
    if (!/^[A-Z]{3}$/.test(expectedPrefix)) return 'Kategorin saknar en giltig trebokstavskod. Rätta kategorin under Administration först.'
    if (prefix !== expectedPrefix) return `Kategorin ska använda prefixet ${expectedPrefix}, inte ${prefix}.`
    const otherCategory = categories.find((category) => category.id !== categoryId && categoryPrefixFor(category.id, categories) === prefix)
    if (otherCategory) return `Prefixet ${prefix} används även av kategorin ${otherCategory.name}. Rätta kategorikoderna under Administration.`
  }
  return ''
}

export const inventoryIntegrityIssues = (items: NamedInventoryTagRecord[], categories: CategoryRecord[]) => {
  const tags = new Map<string, NamedInventoryTagRecord[]>()
  const prefixes = new Map<string, CategoryRecord[]>()
  items.forEach((item) => {
    const tag = normalizeAssetTag(item.assetTag)
    tags.set(tag, [...(tags.get(tag) ?? []), item])
  })
  categories.forEach((category) => {
    const prefix = categoryPrefixFor(category.id, categories)
    prefixes.set(prefix, [...(prefixes.get(prefix) ?? []), category])
  })

  return [
    ...categories.filter((category) => !/^[A-Z]{3}$/.test(categoryPrefixFor(category.id, categories))).map((category) => `${category.name}: ange en kod med exakt tre bokstäver`),
    ...[...prefixes.entries()].filter(([prefix, matchingCategories]) => prefix && matchingCategories.length > 1).map(([prefix, matchingCategories]) => `${prefix} används av kategorierna ${matchingCategories.map((category) => category.name).join(' och ')}`),
    ...items.filter((item) => !ASSET_TAG_PATTERN.test(normalizeAssetTag(item.assetTag))).map((item) => `${item.name}: ${item.assetTag || 'saknar inventarienummer'}`),
    ...[...tags.entries()].filter(([, matchingItems]) => matchingItems.length > 1).map(([tag, matchingItems]) => `${tag} används av ${matchingItems.map((item) => item.name).join(' och ')}`),
    ...items.filter((item) => {
      const expectedPrefix = categoryPrefixFor(item.categoryId, categories)
      return expectedPrefix && ASSET_TAG_PATTERN.test(normalizeAssetTag(item.assetTag)) && !normalizeAssetTag(item.assetTag).startsWith(`${expectedPrefix}-`)
    }).map((item) => `${item.name}: kategorin ska använda ${categoryPrefixFor(item.categoryId, categories)} i stället för ${normalizeAssetTag(item.assetTag).slice(0, 3)}`),
  ]
}
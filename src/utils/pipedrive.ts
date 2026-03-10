export interface PipedriveDeal {
  id: number
  title: string
  value: number | null
  stage_id?: number
  stage_name?: string
  probability?: number | null
  org_name?: string
  hours?: Record<string, number | null>
}

interface DealsResponse {
  data: Array<{
    id: number
    title: string
    value: number | null
    stage_id?: number
    stage_name?: string
    probability?: number | null
    org_id?: { name?: string }
    [key: string]: unknown
  }>
  additional_data?: {
    pagination?: {
      more_items_in_collection?: boolean
      next_start?: number
      limit?: number
      start?: number
    }
  }
}

interface DealField {
  key: string
  name: string
  field_type: string
}

interface DealFieldsResponse {
  data?: DealField[]
}

export async function fetchPipedriveDeals(
  token: string,
  options?: { signal?: AbortSignal; hoursFieldKeys?: Record<string, string | undefined | null> },
): Promise<PipedriveDeal[]> {
  const { signal, hoursFieldKeys = {} } = options ?? {}
  let start = 0
  const pageLimit = 500
  const allDeals: DealsResponse['data'] = []

  // Paginate until Pipedrive reports no more items (max 500 per page)
  while (true) {
    const url = new URL('https://api.pipedrive.com/v1/deals')
    url.searchParams.set('limit', String(pageLimit))
    url.searchParams.set('start', String(start))
    url.searchParams.set('status', 'all_not_deleted')
    url.searchParams.set('api_token', token)

    const response = await fetch(url.toString(), { signal })
    if (!response.ok) {
      throw new Error(`Pipedrive responded with ${response.status}`)
    }

    const body = (await response.json()) as DealsResponse
    const deals = body?.data ?? []
    allDeals.push(...deals)

    const more = body?.additional_data?.pagination?.more_items_in_collection
    const nextStart =
      body?.additional_data?.pagination?.next_start ??
      (body?.additional_data?.pagination?.start ?? 0) + (body?.additional_data?.pagination?.limit ?? pageLimit)
    if (!more) break
    start = nextStart
  }

  return allDeals.map((deal) => ({
    id: deal.id,
    title: deal.title,
    value: deal.value,
    stage_id: deal.stage_id,
    stage_name: deal.stage_name,
    probability: deal.probability,
    org_name: deal.org_id?.name ?? '',
    hours:
      Object.keys(hoursFieldKeys).length === 0
        ? undefined
        : Object.fromEntries(
            Object.entries(hoursFieldKeys).map(([label, key]) => {
              const raw = key ? (deal as any)[key] : undefined
              const num = typeof raw === 'number' ? raw : raw != null ? Number(raw) : null
              return [label, Number.isFinite(num as number) ? (num as number) : null]
            }),
          ),
  }))
}

export async function fetchPipedriveStages(
  token: string,
  signal?: AbortSignal,
): Promise<Record<number, string>> {
  const url = new URL('https://api.pipedrive.com/v1/stages')
  url.searchParams.set('api_token', token)
  const response = await fetch(url.toString(), { signal })
  if (!response.ok) {
    throw new Error(`Pipedrive stages responded with ${response.status}`)
  }
  const json = (await response.json()) as { data?: Array<{ id: number; name: string }> }
  const map: Record<number, string> = {}
  for (const stage of json.data ?? []) {
    map[stage.id] = stage.name
  }
  return map
}

export async function fetchDealFieldKeys(
  token: string,
  signal?: AbortSignal,
): Promise<Record<string, string>> {
  const url = new URL('https://api.pipedrive.com/v1/dealFields')
  url.searchParams.set('api_token', token)
  const response = await fetch(url.toString(), { signal })
  if (!response.ok) {
    throw new Error(`Pipedrive dealFields responded with ${response.status}`)
  }
  const json = (await response.json()) as DealFieldsResponse
  const fields = json.data ?? []

  const wanted: Record<string, string> = {}
  const pick = (label: string, predicate: (f: DealField) => boolean) => {
    const match = fields.find(predicate)
    if (match) wanted[label] = match.key
  }

  pick('fab', (f) => /fab|fabrication/.test(f.name.toLowerCase()) && /hour/.test(f.name.toLowerCase()))
  pick('blast', (f) => /blast/.test(f.name.toLowerCase()) && /hour/.test(f.name.toLowerCase()))
  pick('paint', (f) => /paint/.test(f.name.toLowerCase()) && /hour/.test(f.name.toLowerCase()))
  pick('ship', (f) => /(ship|handling)/.test(f.name.toLowerCase()) && /hour/.test(f.name.toLowerCase()))

  return wanted
}

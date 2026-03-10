export interface PipedriveDeal {
  id: number
  title: string
  value: number | null
  stage_id?: number
  stage_name?: string
  probability?: number | null
  org_name?: string
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
  }>
}

export async function fetchPipedriveDeals(token: string, signal?: AbortSignal): Promise<PipedriveDeal[]> {
  const url = new URL('https://api.pipedrive.com/v1/deals')
  url.searchParams.set('limit', '500')
  url.searchParams.set('api_token', token)

  const response = await fetch(url.toString(), { signal })
  if (!response.ok) {
    throw new Error(`Pipedrive responded with ${response.status}`)
  }

  const body = (await response.json()) as DealsResponse
  const deals = body?.data ?? []

  return deals.map((deal) => ({
    id: deal.id,
    title: deal.title,
    value: deal.value,
    stage_id: deal.stage_id,
    stage_name: deal.stage_name,
    probability: deal.probability,
    org_name: deal.org_id?.name ?? '',
  }))
}

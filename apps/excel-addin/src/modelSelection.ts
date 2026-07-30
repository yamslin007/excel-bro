export interface SelectableModel {
  id: string;
  available: boolean;
}

export function chooseAvailableModel(
  models: SelectableModel[],
  currentModelId: string,
  defaultModelId: string,
  preferredModelId?: string
): string {
  if (
    preferredModelId &&
    models.some(
      (model) => model.id === preferredModelId && model.available
    )
  ) {
    return preferredModelId;
  }
  if (
    models.some(
      (model) => model.id === currentModelId && model.available
    )
  ) {
    return currentModelId;
  }
  return defaultModelId;
}

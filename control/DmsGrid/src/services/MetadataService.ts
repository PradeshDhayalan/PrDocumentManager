import { DataverseClient } from './DataverseClient';
export interface AttributeMetadata {
  LogicalName: string;
  AttributeType: string;
  DisplayName: { UserLocalizedLabel?: { Label: string } };
  RequiredLevel: { Value: string };
  MaxLength?: number;
  IsValidForUpdate: boolean;
  OptionSet?: {
    Options?: {
      Value: number;
      Label: { UserLocalizedLabel?: { Label: string } };
      Color?: string;
    }[];
  };
  Targets?: string[];
}
// TODO(SPIKE-S5): map real option labels, date formats, security and lookup targets here.
export function mapMetadata(attributes: AttributeMetadata[]): AttributeMetadata[] {
  return attributes.filter((attribute) => !!attribute.LogicalName);
}
export class MetadataService {
  private cached?: Promise<AttributeMetadata[]>;
  constructor(private client: DataverseClient) {}
  load(): Promise<AttributeMetadata[]> {
    if (!this.cached)
      this.cached = this.client
        .get<{ Attributes: AttributeMetadata[] }>(
          "EntityDefinitions(LogicalName='dms_document')?$expand=Attributes",
        )
        .then((value) => mapMetadata(value.Attributes))
        .catch((error) => {
          this.cached = undefined;
          throw error;
        });
    return this.cached;
  }
  async lookup(search: string, signal?: AbortSignal): Promise<{ id: string; name: string }[]> {
    const filter = `contains(fullname,'${search.replace(/'/g, "''")}')`;
    const page = await this.client.getPage<{ systemuserid: string; fullname: string }>(
      `systemusers?$filter=${encodeURIComponent(filter)}&$top=25&$orderby=fullname asc`,
      signal,
    );
    return page.value.map((row) => ({ id: row.systemuserid, name: row.fullname }));
  }
}

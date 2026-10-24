import { WebPartContext } from '@microsoft/sp-webpart-base'
import { spfi, SPFI, SPFx, PrincipalSource, PrincipalType } from '@pnp/sp'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
import '@pnp/sp/fields'
import '@pnp/sp/content-types'
import '@pnp/sp/security'
import '@pnp/sp/profiles'
import '@pnp/sp/site-users/web'
import { PermissionKind } from '@pnp/sp/security'
import { IFieldInfo } from '@pnp/sp/fields'
import strings from 'AnnouncementStrings'
import { IListField } from '../components/types'
import { SUPPORTED_FORM_FIELD_TYPES } from '../constants'
import { FALLBACK_FIELDS } from './fallbackFields'

/**
 * A person or group returned by the people picker search
 */
export interface IPersonOption {
  /**
   * Claims key (login name) used with `ensureUser`
   */
  key: string
  text: string
  email?: string
  isGroup: boolean
}

/**
 * Fields from the base content type (Item) that are still useful to show in the popover
 */
const BASE_TYPE_FIELDS_TO_KEEP = ['Title', 'Author', 'Editor', 'Created', 'Modified']

/**
 * Reads an attribute from a field's `SchemaXml`
 */
function schemaAttribute(schemaXml: string, name: string): string | undefined {
  const match = new RegExp(`\s${name}="([^"]*)"`, 'i').exec(schemaXml ?? '')
  return match ? match[1] : undefined
}

/**
 * Maps `SP.Field` to `IListField`
 */
export function mapField(info: IFieldInfo): IListField {
  const choices: any = info.Choices
  return {
    internalName: info.InternalName,
    entityPropertyName: info.EntityPropertyName || info.InternalName,
    title: info.Title,
    description: info.Description ?? '',
    type: info.TypeAsString,
    required: info.Required === true,
    readOnly: info.ReadOnlyField === true,
    hidden: info.Hidden === true,
    fromBaseType: info.FromBaseType === true,
    choices: Array.isArray(choices) ? choices : choices?.results,
    defaultValue: info.DefaultValue ?? undefined,
    richText: /^true$/i.test(schemaAttribute(info.SchemaXml, 'RichText') ?? ''),
    dateOnly:
      info.TypeAsString === 'DateTime' &&
      !/^DateTime$/i.test(schemaAttribute(info.SchemaXml, 'Format') ?? ''),
    allowGroups: /^PeopleAndGroups$/i.test(
      schemaAttribute(info.SchemaXml, 'UserSelectionMode') ?? ''
    ),
    showInNewForm: !/^false$/i.test(schemaAttribute(info.SchemaXml, 'ShowInNewForm') ?? '')
  }
}

/**
 * Whether the field can be edited in the create form
 */
export function isFormField(field: IListField): boolean {
  return (
    !field.hidden &&
    !field.readOnly &&
    field.showInNewForm &&
    (!field.fromBaseType || field.internalName === 'Title') &&
    SUPPORTED_FORM_FIELD_TYPES.includes(field.type)
  )
}

/**
 * Service for the announcement list: schema, permissions, people search and create.
 */
export class AnnouncementListService {
  private _sp: SPFI
  private _fields: Promise<IListField[]>

  constructor(
    context: WebPartContext,
    private _listTitle: string = strings.AnnouncementsListName
  ) {
    this._sp = spfi().using(SPFx(context))
  }

  public get sp(): SPFI {
    return this._sp
  }

  /**
   * The announcement list
   */
  public get list() {
    return this._sp.web.lists.getByTitle(this._listTitle)
  }

  /**
   * Fields of the list's default content type (hidden fields and most base type
   * fields left out), in content type order. Memoized. Falls back to the fields
   * of the template if the schema cannot be read.
   */
  public getFields(): Promise<IListField[]> {
    if (!this._fields) {
      this._fields = this._loadFields().catch(() => FALLBACK_FIELDS)
    }
    return this._fields
  }

  private async _loadFields(): Promise<IListField[]> {
    const contentTypes = await this.list.contentTypes.select('StringId', 'Name', 'Hidden')()
    const defaultContentType =
      contentTypes.find(
        (ct) =>
          ct.StringId.indexOf('0x01') === 0 && ct.StringId.indexOf('0x0120') !== 0 && !ct.Hidden
      ) ?? contentTypes[0]
    const infos: IFieldInfo[] = defaultContentType
      ? await this.list.contentTypes.getById(defaultContentType.StringId).fields()
      : await this.list.fields()
    return infos
      .filter(
        (info) =>
          !info.Hidden &&
          (!info.FromBaseType || BASE_TYPE_FIELDS_TO_KEEP.includes(info.InternalName))
      )
      .map(mapField)
  }

  /**
   * Whether the current user may add items to the list
   */
  public canCreateItems(): Promise<boolean> {
    return this.list.currentUserHasPermissions(PermissionKind.AddListItems).catch(() => false)
  }

  /**
   * Adds an item to the list
   *
   * @param values Field values keyed by internal name (user fields as `<Name>Id`)
   */
  public async addItem(values: Record<string, any>): Promise<void> {
    await this.list.items.add(values)
  }

  /**
   * Searches for people (and optionally groups) using the people picker API
   *
   * @param query Search text
   * @param allowGroups Include security groups and SharePoint groups
   */
  public async searchPeople(query: string, allowGroups: boolean): Promise<IPersonOption[]> {
    if (!query || query.trim().length < 2) return []
    const results = await this._sp.profiles.clientPeoplePickerSearchUser({
      QueryString: query.trim(),
      MaximumEntitySuggestions: 10,
      AllowEmailAddresses: true,
      PrincipalType: allowGroups ? PrincipalType.All : PrincipalType.User,
      PrincipalSource: PrincipalSource.All
    })
    return results.map((result) => ({
      key: result.Key,
      text: result.DisplayText,
      email: result.EntityData?.Email,
      isGroup: result.EntityType !== 'User'
    }))
  }

  /**
   * Ensures the user/group exists on the site and returns its ID
   *
   * @param key Claims key (login name)
   */
  public async ensureUserId(key: string): Promise<number> {
    const user = await this._sp.web.ensureUser(key)
    return user.Id
  }
}

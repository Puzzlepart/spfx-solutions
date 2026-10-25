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
 * Settings that can override the field definition from the content type: the
 * list field (list-level schema) and the content type's field link
 * (hidden/required as set under list settings).
 */
export interface IFieldOverrides {
  listSchemaXml?: string
  linkHidden?: boolean
  linkRequired?: boolean
}

/**
 * Maps `SP.Field` to `IListField`
 *
 * @param info Field from the content type
 * @param overrides List field schema and field link settings
 */
export function mapField(info: IFieldInfo, overrides: IFieldOverrides = {}): IListField {
  const choices: any = info.Choices
  // List schema first: list-level settings are what the list forms use
  const schemas = [overrides.listSchemaXml, info.SchemaXml].filter(Boolean)
  const attribute = (name: string) =>
    schemas.map((schema) => schemaAttribute(schema, name)).filter((value) => value !== undefined)
  const first = (name: string) => attribute(name)[0] ?? ''
  const anyTrue = (name: string) => attribute(name).some((value) => /^true$/i.test(value))
  const anyFalse = (name: string) => attribute(name).some((value) => /^false$/i.test(value))
  return {
    internalName: info.InternalName,
    entityPropertyName: info.EntityPropertyName || info.InternalName,
    title: info.Title,
    description: info.Description ?? '',
    type: info.TypeAsString,
    required: overrides.linkRequired ?? info.Required === true,
    readOnly: info.ReadOnlyField === true,
    hidden: info.Hidden === true || overrides.linkHidden === true || anyTrue('Hidden'),
    fromBaseType: info.FromBaseType === true,
    choices: Array.isArray(choices) ? choices : choices?.results,
    defaultValue: info.DefaultValue ?? undefined,
    richText: /^true$/i.test(first('RichText')),
    dateOnly: info.TypeAsString === 'DateTime' && !/^DateTime$/i.test(first('Format')),
    allowGroups: /^PeopleAndGroups$/i.test(first('UserSelectionMode')),
    showInNewForm: !anyFalse('ShowInNewForm')
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
    let fields: IListField[]
    if (defaultContentType) {
      const contentType = this.list.contentTypes.getById(defaultContentType.StringId)
      const [infos, links, listFields] = await Promise.all([
        contentType.fields(),
        contentType.fieldLinks(),
        this.list.fields()
      ])
      const listFieldById = new Map(listFields.map((field) => [field.Id, field]))
      const linkById = new Map(links.map((link) => [link.Id, link]))
      fields = infos.map((info) =>
        mapField(info, {
          listSchemaXml: listFieldById.get(info.Id)?.SchemaXml,
          linkHidden: linkById.get(info.Id)?.Hidden,
          linkRequired: linkById.get(info.Id)?.Required
        })
      )
    } else {
      fields = (await this.list.fields()).map((info) => mapField(info))
    }
    return fields.filter(
      (field) =>
        !field.hidden &&
        (!field.fromBaseType || BASE_TYPE_FIELDS_TO_KEEP.includes(field.internalName))
    )
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

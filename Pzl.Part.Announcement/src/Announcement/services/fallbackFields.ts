import strings from 'AnnouncementStrings'
import { DEFAULT_SEVERITY, DEFAULT_STATUS_VALUE, SEVERITY_VALUES, STATUS_VALUES } from '../model'
import { IListField } from '../components/types'

const field = (
  internalName: string,
  type: string,
  extra: Partial<IListField> = {}
): IListField => ({
  internalName,
  entityPropertyName: internalName,
  title: strings.FieldTitles[internalName] ?? internalName,
  description: '',
  type,
  required: false,
  readOnly: false,
  hidden: false,
  fromBaseType: false,
  richText: false,
  dateOnly: true,
  allowGroups: false,
  showInNewForm: true,
  ...extra
})

/**
 * Fields of the list as provisioned by `Templates/Announcement.xml`. Used only
 * when the list schema cannot be read from SharePoint. The Teams fields of the
 * notification add-on (Pzl.Ext.Notification) are left out.
 */
export const FALLBACK_FIELDS: IListField[] = [
  field('Title', 'Text', { required: true, fromBaseType: true }),
  field('PzlSeverity', 'Choice', {
    required: true,
    choices: [SEVERITY_VALUES.info, SEVERITY_VALUES.warning, SEVERITY_VALUES.critical],
    defaultValue: SEVERITY_VALUES[DEFAULT_SEVERITY]
  }),
  field('PzlStatus', 'Choice', {
    required: true,
    choices: [STATUS_VALUES.planned, STATUS_VALUES.ongoing, STATUS_VALUES.resolved],
    defaultValue: DEFAULT_STATUS_VALUE
  }),
  field('PzlContent', 'Note', { required: true, richText: true }),
  field('PzlStartDate', 'DateTime', { required: true }),
  field('PzlEndDate', 'DateTime', { required: true }),
  field('PzlAffectedSystems', 'Text'),
  field('PzlConsequences', 'Note'),
  field('PzlResponsible', 'User'),
  field('PzlPinned', 'Boolean'),
  field('PzlShowAsBanner', 'Boolean')
]

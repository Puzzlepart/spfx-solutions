import strings from 'AnnouncementStrings'
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
 * when the list schema cannot be read from SharePoint.
 */
export const FALLBACK_FIELDS: IListField[] = [
  field('Title', 'Text', { required: true, fromBaseType: true }),
  field('PzlSeverity', 'Choice', {
    required: true,
    choices: [
      strings.Severity.Success,
      strings.Severity.Warning,
      strings.Severity.Info,
      strings.Severity.Error
    ],
    defaultValue: strings.Severity.Warning
  }),
  field('PzlContent', 'Note', { required: true, richText: true }),
  field('PzlStartDate', 'DateTime', { required: true }),
  field('PzlEndDate', 'DateTime', { required: true }),
  field('PzlAffectedSystems', 'Text'),
  field('PzlConsequences', 'Note'),
  field('PzlResponsible', 'User'),
  field('PzlPinned', 'Boolean'),
  field('PzlNotifyTeams', 'Boolean'),
  field('PzlNotifiedAt', 'DateTime', { readOnly: true, showInNewForm: false })
]

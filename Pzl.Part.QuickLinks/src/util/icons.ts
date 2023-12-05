import {
  AddCircleFilled,
  AddCircleRegular,
  AddFilled,
  AddRegular,
  LockClosedFilled,
  LockClosedRegular,
  SubtractCircleFilled,
  SubtractCircleRegular,
  bundleIcon
} from '@fluentui/react-icons'

/**
 * Icons for AllLinks component
 *
 */
export const Icons = {
  Add: bundleIcon(AddFilled, AddRegular),
  AddCircle: bundleIcon(AddCircleFilled, AddCircleRegular),
  SubtractCircle: bundleIcon(SubtractCircleFilled, SubtractCircleRegular),
  Lock: bundleIcon(LockClosedFilled, LockClosedRegular)
}

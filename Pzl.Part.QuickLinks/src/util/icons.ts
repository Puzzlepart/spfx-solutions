import {
  Add12Filled,
  Add12Regular,
  Add20Filled,
  Add20Regular,
  Subtract12Filled,
  Subtract12Regular,
  LockClosedFilled,
  LockClosedRegular,
  bundleIcon
} from '@fluentui/react-icons'

/**
 * Icons for AllLinks component
 */
export const Icons = {
  Add20: bundleIcon(Add20Filled, Add20Regular),
  Add: bundleIcon(Add12Filled, Add12Regular),
  Subtract: bundleIcon(Subtract12Filled, Subtract12Regular),
  Lock: bundleIcon(LockClosedFilled, LockClosedRegular)
}

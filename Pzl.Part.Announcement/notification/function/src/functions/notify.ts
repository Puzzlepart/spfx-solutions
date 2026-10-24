import { app, InvocationContext, Timer } from '@azure/functions'
import { getSettings } from '../config'
import { NotificationState } from '../state'
import { processAnnouncements } from '../notifications/announcements'
import { getSiteConfigurations } from '../notifications/configuration'
import { processNews } from '../notifications/news'
import { RecipientResolver } from '../notifications/recipients'

/**
 * Timer-triggered entry point. Runs every minute by default (`TIMER_SCHEDULE`).
 * Timer triggers are singletons, so two runs never overlap within one function app.
 */
export async function notify(timer: Timer, context: InvocationContext): Promise<void> {
  const settings = getSettings()
  if (timer.isPastDue) context.warn('Timer is past due, running now')
  if (settings.dryRun) context.warn('DRY_RUN is enabled: nothing is sent or written')

  const sites = await getSiteConfigurations(settings)
  context.log(`${sites.length} active site configuration(s) in ${settings.configSiteUrl}`)

  const recipients = new RecipientResolver(settings, context)
  const state = new NotificationState(settings, context)

  for (const site of sites) {
    if (!site.defaultGroupIds.length) {
      context.warn(
        `${site.title}: no default recipient groups configured, items without audience are skipped`
      )
    }
    if (site.notifyAnnouncements) {
      try {
        await processAnnouncements(site, settings, recipients, context)
      } catch (error) {
        context.error(`${site.title}: announcements failed: ${error}`)
      }
    }
    if (site.notifyNews) {
      try {
        await processNews(site, settings, recipients, state, context)
      } catch (error) {
        context.error(`${site.title}: news failed: ${error}`)
      }
    }
  }
}

app.timer('notify', {
  schedule: process.env.TIMER_SCHEDULE || '0 * * * * *',
  runOnStartup: false,
  handler: notify
})

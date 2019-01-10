package org.exoplatform.extension.exchange.service;

import java.util.*;
import java.util.concurrent.*;

import org.apache.commons.lang.StringUtils;
import org.picocontainer.Startable;

import com.google.common.util.concurrent.ThreadFactoryBuilder;

import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.commons.utils.CommonsUtils;
import org.exoplatform.container.xml.InitParams;
import org.exoplatform.extension.exchange.task.ExchangeIntegrationTask;
import org.exoplatform.extension.exchange.task.UserIntegrationFacade;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.security.*;

/**
 * Service to listen to exchange events. Used to synchronize eXo User calendar
 * state with exchange User calendar in real time. This service is used by two
 * modules: LoginModule to register User subscription to exchange events and a
 * LogoutListener that will be used to
 * 
 * @author Boubaker KHANFIR
 */
public class SynchronizationService implements Startable {

  private static final Log                      LOG                                       =
                                                    ExoLogger.getLogger(SynchronizationService.class);

  private static final String                   EXCHANGE_SERVER_URL_PARAM_NAME            = "exchange.ews.url";

  private static final String                   EXCHANGE_DOMAIN_PARAM_NAME                = "exchange.domain";

  private static final String                   EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME    = "exchange.scheduler.delay";

  private static final int                      EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM = 5;

  private static final String                   EXCHANGE_SYNCHRONIZE_ALL                  = "exchange.synchronize.all.folders";

  private static final String                   EXCHANGE_MAX_DAYS                         = "exchange.synchronize.max.days";

  private static final String                   EXCHANGE_DELETE_CALENDAR_ON_UNSYNC        = "exchange.delete.calendar.on.unsync";

  private int                                   schedulerDelayInSeconds                   =
                                                                        EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM;

  private final ScheduledExecutorService        scheduledExecutor;

  private final Map<String, ScheduledFuture<?>> futures                                   = new HashMap<>();

  private final Map<String, Runnable>           runnables                                 = new HashMap<>();

  private final ExoDataStorageService           exoStorageService;

  private final ExchangeDataStorageService      exchangeStorageService;

  private final CorrespondenceService           correspondenceService;

  private OrganizationService                   organizationService;

  private CalendarService                       calendarService;

  private IdentityRegistry                      identityRegistry;

  private String                                exchangeServerURL;

  private String                                exchangeDomain;

  private boolean                               synchronizeAllExchangeFolders             = false;

  private boolean                               deleteExoCalendarOnUnsync                 = false;

  private int                                   maxFirstSynchronizationDays               = 365;

  public SynchronizationService(ExoDataStorageService exoStorageService,
                                ExchangeDataStorageService exchangeStorageService,
                                CorrespondenceService correspondenceService,
                                InitParams params) {
    this.exoStorageService = exoStorageService;
    this.exchangeStorageService = exchangeStorageService;
    this.correspondenceService = correspondenceService;

    ThreadFactory namedThreadFactory = new ThreadFactoryBuilder().setNameFormat("ExchangeSynchronization-%d").build();
    this.scheduledExecutor = Executors.newScheduledThreadPool(10, namedThreadFactory);

    if (params.containsKey(EXCHANGE_SERVER_URL_PARAM_NAME)
        && !params.getValueParam(EXCHANGE_SERVER_URL_PARAM_NAME).getValue().isEmpty()) {
      this.exchangeServerURL = params.getValueParam(EXCHANGE_SERVER_URL_PARAM_NAME).getValue();
    } else {
      LOG.warn("Echange Synchronization Service: Default MS Exchange server URL (init-param " + EXCHANGE_SERVER_URL_PARAM_NAME
          + ") is not set.");
    }
    if (params.containsKey(EXCHANGE_DOMAIN_PARAM_NAME)
        && !params.getValueParam(EXCHANGE_DOMAIN_PARAM_NAME).getValue().isEmpty()) {
      this.exchangeDomain = params.getValueParam(EXCHANGE_DOMAIN_PARAM_NAME).getValue();
    } else {
      LOG.warn("Echange Synchronization Service: Default MS Exchange domain name (init-param " + EXCHANGE_DOMAIN_PARAM_NAME
          + ") is not set.");
    }
    if (params.containsKey(EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME)) {
      String schedulerDelayInSecondsString = params.getValueParam(EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME).getValue();
      this.schedulerDelayInSeconds = Integer.valueOf(schedulerDelayInSecondsString);
    } else {
      LOG.warn("Echange Synchronization Service: Check Period in seconds (init-param {}) is not set. Default will be used: {} seconds.",
               EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME,
               EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM);
      this.schedulerDelayInSeconds = EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM;
    }
    if (schedulerDelayInSeconds < EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM) {
      LOG.warn("Echange Synchronization Service: Check Period in seconds (init-param {}) is set under {} seconds. Default will be used: {} seconds.",
               EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME,
               EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM,
               EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM);
      this.schedulerDelayInSeconds = EXCHANGE_LISTENER_SCHEDULER_DELAY_MINIMUM;
    }
    if (params.containsKey(EXCHANGE_DELETE_CALENDAR_ON_UNSYNC)) {
      String deleteExoCalendarOnUnsyncString = params.getValueParam(EXCHANGE_DELETE_CALENDAR_ON_UNSYNC).getValue();
      if (deleteExoCalendarOnUnsyncString != null && deleteExoCalendarOnUnsyncString.equals("true")) {
        this.deleteExoCalendarOnUnsync = true;
      }
    }
    if (params.containsKey(EXCHANGE_SYNCHRONIZE_ALL)) {
      String exchangeSynchronizeAllString = params.getValueParam(EXCHANGE_SYNCHRONIZE_ALL).getValue();
      if (exchangeSynchronizeAllString != null && exchangeSynchronizeAllString.equals("true")) {
        this.synchronizeAllExchangeFolders = true;
      }
    }

    if (params.containsKey(EXCHANGE_MAX_DAYS)) {
      String exchangeMaxDaysString = params.getValueParam(EXCHANGE_MAX_DAYS).getValue();
      if (StringUtils.isNotBlank(exchangeMaxDaysString)) {
        this.maxFirstSynchronizationDays = Integer.parseInt(exchangeMaxDaysString);
      }
    }
  }

  @Override
  public void start() {
    LOG.info("Echange Synchronization Service: Successfully started.");
  }

  @Override
  public void stop() {
    scheduledExecutor.shutdownNow();
  }

  public String getExchangeDomain() {
    return exchangeDomain;
  }

  public String getExchangeServerURL() {
    return exchangeServerURL;
  }

  /**
   * Register User with Exchange services.
   * 
   * @param username
   * @param password
   */
  public void userLoggedIn(final String username, final String password) throws Exception {
    String exchangeStoredUsername =
                                  UserIntegrationFacade.getUserArrtibute(getOrganizationService(),
                                                                         username,
                                                                         UserIntegrationFacade.USER_EXCHANGE_USERNAME_ATTRIBUTE);
    if (StringUtils.isNotBlank(exchangeStoredUsername)) {
      String exchangeStoredServerName =
                                      UserIntegrationFacade.getUserArrtibute(getOrganizationService(),
                                                                             username,
                                                                             UserIntegrationFacade.USER_EXCHANGE_SERVER_URL_ATTRIBUTE);
      String exchangeStoredDomainName =
                                      UserIntegrationFacade.getUserArrtibute(getOrganizationService(),
                                                                             username,
                                                                             UserIntegrationFacade.USER_EXCHANGE_SERVER_DOMAIN_ATTRIBUTE);
      String exchangeStoredPassword =
                                    UserIntegrationFacade.getUserArrtibute(getOrganizationService(),
                                                                           username,
                                                                           UserIntegrationFacade.USER_EXCHANGE_PASSWORD_ATTRIBUTE);
      startExchangeSynchronizationTask(username,
                                       exchangeStoredUsername,
                                       exchangeStoredPassword,
                                       exchangeStoredDomainName,
                                       exchangeStoredServerName);
    } else if (StringUtils.isNotBlank(exchangeDomain) && StringUtils.isNotBlank(exchangeServerURL)) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Exchange Synchronization Service: User '" + username
            + "' have not yet set parameters, use default Exchange server settings.");
      }
      startExchangeSynchronizationTask(username, username, password, exchangeDomain, exchangeServerURL);
    } else {
      LOG.warn("Exchange Service is unvailable, please set parameters.");
    }
  }

  /**
   * Register User with Exchange services.
   * 
   * @param username
   * @param password
   * @param exchangeDomain
   * @param exchangeServerURL
   */
  public void startExchangeSynchronizationTask(final String username,
                                               String exchangeUsername,
                                               final String password,
                                               String exchangeDomain,
                                               String exchangeServerURL) {
    try {
      if (StringUtils.isBlank(password) || StringUtils.isBlank(exchangeUsername) || StringUtils.isBlank(exchangeServerURL)) {
        return;
      }
      exchangeUsername = exchangeUsername.trim();
      Identity identity = getIdentityRegistry().getIdentity(username);
      if (identity == null || identity.getUserId().equals(IdentityConstants.ANONIM)) {
        throw new IllegalStateException("Identity of user '" + username + "' not found.");
      }

      // Close other tasks if already exists, this can happens when user is
      // still logged in in other browser
      closeTaskIfExists(username);

      // Scheduled task: listen the changes made on MS Exchange Calendar
      Runnable schedulerCommand = new ExchangeIntegrationTask(getCalendarService(),
                                                              exoStorageService,
                                                              exchangeStorageService,
                                                              correspondenceService,
                                                              identity,
                                                              exchangeUsername,
                                                              password,
                                                              exchangeDomain,
                                                              exchangeServerURL,
                                                              synchronizeAllExchangeFolders,
                                                              deleteExoCalendarOnUnsync,
                                                              maxFirstSynchronizationDays);

      ScheduledFuture<?> future = scheduledExecutor.scheduleWithFixedDelay(schedulerCommand,
                                                                           10,
                                                                           schedulerDelayInSeconds,
                                                                           TimeUnit.SECONDS);

      // Add future task to the map to destroy thread when the user logout
      futures.put(username, future);
      runnables.put(username, schedulerCommand);

      LOG.info("User '" + username + "' logged in, exchange synchronization task started.");
    } catch (Exception e) {
      LOG.warn("Exchange integration error for user '" + username + "' : ", e);
      if (LOG.isTraceEnabled() || LOG.isDebugEnabled()) {
        LOG.trace("Error while initializing user integration with exchange: ", e);
      }
    }
  }

  /**
   * Unregister User from Exchange services.
   * 
   * @param username
   */
  public void userLoggedOut(String username) {
    ConversationRegistry conversationRegistry = CommonsUtils.getService(ConversationRegistry.class);
    List<StateKey> stateKeys = conversationRegistry.getStateKeys(username);
    if (stateKeys == null || stateKeys.isEmpty()) {
      closeTaskIfExists(username);
    }
  }

  /**
   * Forces the execution of synchronization
   * 
   * @param username
   */
  public void synchronize(String username) {
    Runnable command = runnables.get(username);
    if (command != null) {
      command.run();
    }
  }

  private void closeTaskIfExists(String username) {
    ScheduledFuture<?> future = futures.remove(username);
    if (future != null) {
      future.cancel(true);
      UserIntegrationFacade integrationService = UserIntegrationFacade.getInstance(username);
      if (integrationService != null) {
        try {
          integrationService.removeInstance();
        } catch (Throwable e) {
          // Nothing to do, just log this.
          LOG.error(e);
        }
      }
      LOG.info("Exchange synchronization task stopped for User '" + username + "'.");
    }
  }

  public CalendarService getCalendarService() {
    if (calendarService == null) {
      calendarService = CommonsUtils.getService(CalendarService.class);
    }
    return calendarService;
  }

  public OrganizationService getOrganizationService() {
    if (organizationService == null) {
      organizationService = CommonsUtils.getService(OrganizationService.class);
    }
    return organizationService;
  }

  public IdentityRegistry getIdentityRegistry() {
    if (identityRegistry == null) {
      identityRegistry = CommonsUtils.getService(IdentityRegistry.class);
    }
    return identityRegistry;
  }
}

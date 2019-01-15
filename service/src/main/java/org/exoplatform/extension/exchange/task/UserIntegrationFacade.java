package org.exoplatform.extension.exchange.task;

import java.io.*;
import java.security.KeyStore;
import java.util.*;

import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.jcr.Node;

import org.apache.commons.lang3.StringUtils;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.commons.api.settings.SettingService;
import org.exoplatform.commons.api.settings.SettingValue;
import org.exoplatform.commons.api.settings.data.Context;
import org.exoplatform.commons.api.settings.data.Scope;
import org.exoplatform.commons.utils.CommonsUtils;
import org.exoplatform.commons.utils.PropertyManager;
import org.exoplatform.container.PortalContainer;
import org.exoplatform.container.component.ComponentRequestLifecycle;
import org.exoplatform.extension.exchange.service.*;
import org.exoplatform.extension.exchange.service.util.CalendarConverterUtils;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.organization.UserProfile;
import org.exoplatform.web.security.codec.AbstractCodec;
import org.exoplatform.web.security.codec.AbstractCodecBuilder;
import org.exoplatform.web.security.security.TokenServiceInitializationException;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.ItemTraversal;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.enumeration.service.SyncFolderItemsScope;
import microsoft.exchange.webservices.data.core.enumeration.sync.ChangeType;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.exception.service.remote.ServiceRemoteException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.notification.ItemEvent;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;
import microsoft.exchange.webservices.data.sync.ChangeCollection;
import microsoft.exchange.webservices.data.sync.ItemChange;

/**
 * @author Boubaker KHANFIR
 */
@SuppressWarnings("all")
public class UserIntegrationFacade {

  private static final String                             GTN_KEY                               = "gtnKey";

  private static final String                             GTN_STORE_PASS                        = "gtnStorePass";

  public static final String                              USER_EXCHANGE_SERVER_URL_ATTRIBUTE    = "exchange.server.url";

  public static final String                              USER_EXCHANGE_SERVER_DOMAIN_ATTRIBUTE = "exchange.server.domain";

  public static final String                              USER_EXCHANGE_USERNAME_ATTRIBUTE      = "exchange.username";

  public static final String                              USER_EXCHANGE_PASSWORD_ATTRIBUTE      = "exchange.password";

  public static final Context                             USER_EXCHANGE_CONTEXT                 =
                                                                                Context.GLOBAL.id("ADDONS_EXCHANGE_CALENDAR");

  public static final Scope                               USER_EXCHANGE_SCOPE                   = Scope.APPLICATION;

  public static final String                              USER_EXCHANGE_SYNC_STATE_KEY          = "ADDONS_EXCHANGE_SYNC_STATE";

  private final static Log                                LOG                                   =
                                                              ExoLogger.getLogger(UserIntegrationFacade.class);

  private static final String                             USER_EXO_HANDLED_ATTRIBUTE            = "exo.check.date";

  private static final Map<String, UserIntegrationFacade> instances                             =
                                                                    new HashMap<String, UserIntegrationFacade>();

  private static AbstractCodec                            codec;

  private final String                                    username;

  private final int                                       maxFirstSynchronizationDays;

  private final ExchangeService                           service;

  private final ExoDataStorageService                     exoStorageService;

  private final ExchangeDataStorageService                exchangeStorageService;

  private final CorrespondenceService                     correspondenceService;

  private final CalendarService                           calendarService;

  private SettingService                                  settingService;

  private boolean                                         synchIsCurrentlyRunning               = false;

  private Date                                            firstSynchronizationUntilDate;

  public UserIntegrationFacade(CalendarService calendarService,
                               ExoDataStorageService exoStorageService,
                               ExchangeDataStorageService exchangeStorageService,
                               CorrespondenceService correspondenceService,
                               ExchangeService service,
                               String username,
                               int maxFirstSynchronizationDays) {
    this.calendarService = calendarService;
    this.exoStorageService = exoStorageService;
    this.exchangeStorageService = exchangeStorageService;
    this.correspondenceService = correspondenceService;
    this.service = service;
    this.username = username;
    this.maxFirstSynchronizationDays = maxFirstSynchronizationDays;

    // Set corresponding service to each username.
    instances.put(username, this);

    java.util.Calendar untilCalendarDate = java.util.Calendar.getInstance();
    untilCalendarDate.add(java.util.Calendar.DATE, -maxFirstSynchronizationDays);
    firstSynchronizationUntilDate = untilCalendarDate.getTime();
  }

  /**
   * Gets user exchange instance service.
   * 
   * @param username
   * @return
   */
  public static UserIntegrationFacade getInstance(String username) {
    return instances.get(username);
  }

  /**
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(FolderId folderId) throws Exception {
    return exchangeStorageService.getExchangeCalendar(service, folderId);
  }

  /**
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * 
   * @param folderId
   * @throws Exception
   * @return List of event IDs
   */
  public List<String> synchronizeFullCalendar(FolderId folderId) throws Exception {
    List<String> updatedExoEventIds = new ArrayList<String>();
    CalendarFolder folder = exchangeStorageService.getExchangeCalendar(service, folderId);

    Calendar exoCalendar = exoStorageService.getUserCalendar(username, folder.getId().getUniqueId());
    boolean isNewCalendar = exoCalendar == null;

    String syncState = getSynchState(folderId);
    if (syncState == null || isNewCalendar) {
      LOG.debug("Start full exchange calendar synchronization for user '{}' exchange folder calendar {} until date {}",
                username,
                folderId.getFolderName() == null ? folderId.getUniqueId() : folderId.getFolderName(),
                firstSynchronizationUntilDate);

      Date lastSynchronizedDate = null;

      // Create Calendar if not present
      exoStorageService.getOrCreateUserCalendar(username, folder);

      int offset = 0;
      int pageSize = 30;

      ItemView view = new ItemView(pageSize);
      view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
      view.getOrderBy().add(AppointmentSchema.Start, SortDirection.Descending);
      view.setTraversal(ItemTraversal.Shallow);

      // Synchronize all events
      ChangeCollection<ItemChange> changeCollection = service.syncFolderItems(folderId,
                                                                              PropertySet.FirstClassProperties,
                                                                              null,
                                                                              512,
                                                                              SyncFolderItemsScope.NormalItems,
                                                                              syncState);
      syncState = changeCollection.getSyncState();

      while (true) {
        FindItemsResults<Item> results = null;
        try {
          results = getItems(folderId, view);
        } catch (ServiceRemoteException e) {
          results = getItems(folderId, view);
        }

        try {
          lastSynchronizedDate = synchronizeExchangeAppointments(updatedExoEventIds, results.getItems());
        } catch (Exception e) {
          LOG.error("Error while synchronizing for user '{}' '{}' items from offset '{}'", username, pageSize, offset);
        }

        if (!results.isMoreAvailable()
            || (lastSynchronizedDate != null && firstSynchronizationUntilDate.after(lastSynchronizedDate))) {
          break;
        }

        offset += pageSize;
        view.setOffset(offset);
      }

      LOG.debug("Full exchange calendar synchronization processed successfully for user '{}' for exchange calendar {}, last synchronized event start date: '{}'",
                username,
                folderId.getFolderName() == null ? folderId.getUniqueId() : folderId.getFolderName(),
                lastSynchronizedDate == null ? "NO DATE" : lastSynchronizedDate);
      setSynchState(folderId, syncState);
    } else {
      LOG.debug("Synchronize last modified events since last synchronization for user '{}'", username);

      int countModifiedItems = synchronizeExchangeAppointementsByState(folderId, updatedExoEventIds);

      LOG.debug("First synchronization is finished for user '{}' with {} modified/created events", username, countModifiedItems);
    }

    return updatedExoEventIds;
  }

  private int synchronizeExchangeAppointementsByState(FolderId folderId, List<String> updatedExoEventIds) throws Exception {
    String syncState = getSynchState(folderId);
    ChangeCollection<ItemChange> changeCollection = service.syncFolderItems(folderId,
                                                                            PropertySet.FirstClassProperties,
                                                                            null,
                                                                            512,
                                                                            SyncFolderItemsScope.NormalItems,
                                                                            syncState);
    syncState = changeCollection.getSyncState();
    int countModifiedItems = changeCollection.getCount();

    List<Item> modifiedItems = new ArrayList<>();
    Iterator<ItemChange> changeIterator = changeCollection.iterator();
    while (changeIterator.hasNext()) {
      ItemChange action = (ItemChange) changeIterator.next();
      Item item = action.getItem();
      if (item instanceof Appointment) {
        if (ChangeType.Create.equals(action.getChangeType()) || ChangeType.Update.equals(action.getChangeType())) {
          modifiedItems.add(item);
        } else if (ChangeType.Delete.equals(action.getChangeType())) {
          String itemId = item.getId().getUniqueId();
          checkAndDeleteExoEvent(itemId);
        }
      }
    }

    if (!modifiedItems.isEmpty()) {
      synchronizeExchangeAppointments(updatedExoEventIds, modifiedItems);
    }
    setSynchState(folderId, syncState);
    return countModifiedItems;
  }

  public void checkAndDeleteExoEvent(String itemId) throws Exception {
    String eventId = correspondenceService.getCorrespondingId(username, itemId);
    CalendarEvent calendarEvent = exoStorageService.getEvent(eventId, username);
    if (calendarEvent != null) {
      exoStorageService.deleteEvent(username, calendarEvent);
    }
  }

  private void setSynchState(FolderId folderId, String syncState) {
    getSettingService().set(USER_EXCHANGE_CONTEXT,
                            USER_EXCHANGE_SCOPE.id(folderId.getUniqueId()),
                            USER_EXCHANGE_SYNC_STATE_KEY,
                            SettingValue.create(syncState));
  }

  public String getSynchState(FolderId folderId) {
    SettingValue<?> settingValue = getSettingService().get(USER_EXCHANGE_CONTEXT,
                                                           USER_EXCHANGE_SCOPE.id(folderId.getUniqueId()),
                                                           USER_EXCHANGE_SYNC_STATE_KEY);
    return settingValue == null || settingValue.getValue() == null ? null : settingValue.getValue().toString();
  }

  public SettingService getSettingService() {
    if (settingService == null) {
      settingService = CommonsUtils.getService(SettingService.class);
    }
    return settingService;
  }

  /**
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * The check is done for events modified since 'lastSyncDate'.
   * 
   * @param folderId
   * @param exoLastSyncDate
   * @param updatedExoEventIDs
   * @throws Exception
   */
  public void synchronizeModificationsOfCalendar(FolderId folderId,
                                                 Date exoLastSyncDate,
                                                 List<String> updatedExoEventIDs) throws Exception {
    synchronizeExchangeAppointementsByState(folderId, updatedExoEventIDs);
    synchronizeExoEventsByModificationDate(folderId, updatedExoEventIDs, exoLastSyncDate);
  }

  /**
   * Gets list of personnal Exchange Calendars.
   * 
   * @return list of FolderId
   * @throws Exception
   */
  public List<FolderId> getAllExchangeCalendars() throws Exception {
    return exchangeStorageService.getAllExchangeCalendars(service);
  }

  /**
   * Checks if eXo associated Calendar is present.
   * 
   * @param id
   * @return true if present.
   * @throws Exception
   */
  public boolean isCalendarSynchronizedWithExchange(String id) throws Exception {
    return StringUtils.isNotBlank(id) && correspondenceService.getCorrespondingId(username, id) != null;
  }

  /**
   * Checks if eXo associated Calendar is present.
   * 
   * @param folderId
   * @return true if present.
   * @throws Exception
   */
  public boolean isCalendarPresentInExo(FolderId folderId) throws Exception {
    return exoStorageService.getUserCalendar(username, folderId.getUniqueId()) != null;
  }

  /**
   * Creates or updates or deletes eXo Calendar Event associated to Item, switch
   * state in Exchange.
   * 
   * @param itemEvent
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> createOrUpdateOrDelete(ItemEvent itemEvent) throws Exception {
    List<CalendarEvent> updatedEvents = null;

    Item item = exchangeStorageService.getItem(service, itemEvent.getItemId());
    if (item == null) {
      exoStorageService.deleteEventByAppointmentID(itemEvent.getItemId().getUniqueId(), username);
    } else if (item instanceof Appointment) {
      Appointment appointment = (Appointment) item;
      String eventId = correspondenceService.getCorrespondingId(username, appointment.getId().getUniqueId());
      if (eventId == null) {
        updatedEvents = exoStorageService.createEvent(appointment, username);
      } else {
        updatedEvents = exoStorageService.updateEvent(appointment, username);
      }
    }

    return updatedEvents;
  }

  /**
   * Get corresponding User Calenar from Exchange Folder Id
   * 
   * @param folderId
   * @return
   * @throws Exception
   */
  public Calendar getUserCalendarByExchangeFolderId(FolderId folderId) throws Exception {
    return exoStorageService.getUserCalendar(username, folderId.getUniqueId());
  }

  /**
   * @param eventNode
   * @throws Exception
   */
  public void updateOrCreateExchangeCalendarEvent(Node eventNode) throws Exception {
    CalendarEvent event = exoStorageService.getExoEventByNode(eventNode);
    if (isCalendarSynchronizedWithExchange(event.getCalendarId())) {
      updateOrCreateExchangeCalendarEvent(event);
    }
  }

  /**
   * @param eventId
   * @throws Exception
   */
  public boolean updateOrCreateExchangeCalendarEvent(String eventId) throws Exception {
    CalendarEvent event = ((CalendarServiceImpl) calendarService).getDataStorage().getEvent(username, eventId);
    return updateOrCreateExchangeCalendarEvent(event);
  }

  /**
   * @param event
   * @throws Exception
   */
  public boolean updateOrCreateExchangeCalendarEvent(CalendarEvent event) throws Exception {
    String exoMasterId = null;
    if (event.getIsExceptionOccurrence() != null && event.getIsExceptionOccurrence()) {
      exoMasterId = exoStorageService.getExoEventMasterRecurenceByOriginalUUID(event.getOriginalReference());
      if (exoMasterId == null) {
        LOG.error("No master Id was found for occurence: " + event.getSummary() + " with recurrenceId = "
            + event.getRecurrenceId() + ". The event will not be updated.");
      }
    }
    return exchangeStorageService.updateOrCreateExchangeAppointment(username,
                                                                    service,
                                                                    event,
                                                                    exoMasterId,
                                                                    this::appointmentUpdated);
  }

  /**
   * @param eventId
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeCalendarEvent(String eventId, String calendarId) throws Exception {
    exchangeStorageService.deleteAppointmentByExoEventId(username, service, eventId, calendarId);
  }

  /**
   * Handle Exchange Calendar Deletion by deleting associated eXo Calendar.
   * 
   * @param folderId Exchange Calendar folderId
   * @return
   * @throws Exception
   */
  public boolean deleteExoCalendar(FolderId folderId) throws Exception {
    Folder folder = exchangeStorageService.getExchangeCalendar(service, folderId);
    if (folder != null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Folder was found, but event seems saying that it was deleted.");
      }
      return false;
    }
    return exoStorageService.deleteCalendar(username, folderId.getUniqueId());
  }

  /**
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeCalendar(String calendarId) throws Exception {
    exchangeStorageService.deleteExchangeFolderByCalenarId(username, service, calendarId);
  }

  public void removeInstance() {
    if (LOG.isTraceEnabled()) {
      LOG.trace("Stop Exchange Integration Service for user: " + username);
    }
    instances.remove(username);
  }

  public List<String> synchronizeExchangeFolderState(List<FolderId> calendarFolderIds,
                                                     boolean synchronizeAllExchangeFolders,
                                                     boolean deleteExoCalendarOnUnsync) throws Exception {
    Iterator<FolderId> iterator = calendarFolderIds.iterator();
    while (iterator.hasNext()) {
      FolderId folderId = (FolderId) iterator.next();
      deleteExoCalendarOutOfSync(deleteExoCalendarOnUnsync, iterator, folderId);
    }

    List<String> updatedCalendarEventIds = null;
    // synchronize added Folders
    if (synchronizeAllExchangeFolders) {
      List<FolderId> folderIds = exchangeStorageService.getAllExchangeCalendars(service);
      for (FolderId folderId : folderIds) {
        // Test if not already fully synchronized
        if (!calendarFolderIds.contains(folderId)) {
          // Delete eXo calendar and recreate it
          exoStorageService.deleteCalendar(username, folderId.getUniqueId());

          List<String> tmpUpdatedCalendarEventIds = synchronizeFullCalendar(folderId);
          if (tmpUpdatedCalendarEventIds != null && !tmpUpdatedCalendarEventIds.isEmpty()) {
            if (updatedCalendarEventIds == null) {
              updatedCalendarEventIds = tmpUpdatedCalendarEventIds;
            } else {
              updatedCalendarEventIds.addAll(tmpUpdatedCalendarEventIds);
            }
          }

          calendarFolderIds.add(folderId);
        }
      }
    } else {
      // Checks if synchronized Exchange Folders have been modified
      List<FolderId> synchronizedFolderIds = getSynchronizedExchangeCalendars();
      for (FolderId folderId : synchronizedFolderIds) {
        // Test if not already fully synchronized
        if (!calendarFolderIds.contains(folderId)) {
          // Delete eXo calendar and recreate it
          exoStorageService.deleteCalendar(username, folderId.getUniqueId());

          List<String> tmpUpdatedCalendarEventIds = synchronizeFullCalendar(folderId);
          if (tmpUpdatedCalendarEventIds != null && !tmpUpdatedCalendarEventIds.isEmpty()) {
            if (updatedCalendarEventIds == null) {
              updatedCalendarEventIds = tmpUpdatedCalendarEventIds;
            } else {
              updatedCalendarEventIds.addAll(tmpUpdatedCalendarEventIds);
            }
          }

          calendarFolderIds.add(folderId);
        }
      }
      Iterator<FolderId> folderIdIterator = calendarFolderIds.iterator();
      while (folderIdIterator.hasNext()) {
        FolderId folderId = (FolderId) folderIdIterator.next();
        if (!synchronizedFolderIds.contains(folderId)) {
          folderIdIterator.remove();
        }
      }
    }
    return updatedCalendarEventIds;
  }

  private void deleteExoCalendarOutOfSync(boolean deleteExoCalendarOnUnsync,
                                          Iterator<FolderId> iterator,
                                          FolderId folderId) throws Exception {
    Folder folder = exchangeStorageService.getExchangeCalendar(service, folderId);
    if (folder == null) {
      // Test if the connection is ok, else the exception is thrown because of
      // interrupted connection
      folder =
             exchangeStorageService.getExchangeCalendar(service,
                                                        FolderId.getFolderIdFromWellKnownFolderName(WellKnownFolderName.Calendar));

      if (folder != null) {
        Calendar calendar = exoStorageService.getUserCalendar(username, folderId.getUniqueId());
        if (calendar != null) {
          if (LOG.isTraceEnabled()) {
            LOG.trace("Folder '" + folderId.getUniqueId()
                + "' was deleted from Exchange, stopping synchronization for this folder.");
          }
          if (deleteExoCalendarOnUnsync) {
            exoStorageService.deleteCalendar(username, folderId.getUniqueId());
          } else {
            correspondenceService.deleteCorrespondingId(username, folderId.getUniqueId());
          }
          // Remove FolderId from synchronized folders in Scheduled Job
          iterator.remove();
        }
      }
    }
  }

  public List<FolderId> getSynchronizedExchangeCalendars() throws Exception {
    List<FolderId> folderIds = new ArrayList<FolderId>();
    List<String> folderIdsString = correspondenceService.getSynchronizedExchangeFolderIds(username);
    for (String folderIdString : folderIdsString) {
      folderIds.add(FolderId.getFolderIdFromString(folderIdString));
    }
    return folderIds;
  }

  public void addFolderToSynchronization(String folderIdString) throws Exception {
    String calendarId = CalendarConverterUtils.getCalendarId(folderIdString);
    correspondenceService.setCorrespondingId(username, calendarId, folderIdString);
  }

  public void deleteFolderFromSynchronization(String folderIdString) throws Exception {
    correspondenceService.deleteCorrespondingId(username, folderIdString);
  }

  public synchronized boolean setSynchronizationStarted() {
    return synchIsCurrentlyRunning ? false : (synchIsCurrentlyRunning = true);
  }

  public synchronized void setSynchronizationStopped() {
    synchIsCurrentlyRunning = false;
  }

  private void deleteExoEventsOutOfSynchronization(FolderId folderId) throws Exception {
    List<CalendarEvent> events = exoStorageService.getUserCalendarEvents(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : events) {
      String itemId = correspondenceService.getCorrespondingId(username, calendarEvent.getId());
      if (itemId == null) {
        exoStorageService.deleteEvent(username, calendarEvent);
      } else {
        Item item = exchangeStorageService.getItem(service, itemId);
        if (item == null) {
          exoStorageService.deleteEvent(username, calendarEvent);
        }
      }
    }
  }

  private Date synchronizeExchangeAppointments(List<String> eventIds, Iterable<Item> items) throws Exception,
                                                                                            ServiceLocalException {
    Date lastSynchronizedDate = null;
    for (Item item : items) {
      if (item instanceof Appointment) {
        Appointment appointment = (Appointment) item;
        if (lastSynchronizedDate == null || lastSynchronizedDate.before(appointment.getStart())) {
          lastSynchronizedDate = appointment.getStart();
        }

        List<CalendarEvent> updatedEvents = null;
        try {
          updatedEvents = exoStorageService.createOrUpdateEvent((Appointment) item, username);
        } catch (Exception e) {
          LOG.warn("Error user '{}' create/update exchange item '{}'", username, item.getId().getUniqueId());
        }

        if (updatedEvents != null && !updatedEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : updatedEvents) {
            eventIds.add(calendarEvent.getId());
          }
        }
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
    return lastSynchronizedDate;
  }

  /**
   * Sets exo last check operation date.
   * 
   * @param time
   * @throws Exception
   */
  public void setUserExoLastCheckDate(long time) throws Exception {
    if (LOG.isTraceEnabled()) {
      LOG.trace("Set last time check for modified exo events to '{}.{}' for user {}", new Date(time), time % 1000, username);
    }
    getSettingService().set(USER_EXCHANGE_CONTEXT,
                            USER_EXCHANGE_SCOPE.id(username),
                            USER_EXO_HANDLED_ATTRIBUTE,
                            SettingValue.create(String.valueOf(time)));
  }

  /**
   * Gets exo last check operation date
   * 
   * @return
   * @throws Exception
   */
  public Date getUserExoLastCheckDate() throws Exception {
    SettingValue<?> settingValue = getSettingService().get(USER_EXCHANGE_CONTEXT,
                                                           USER_EXCHANGE_SCOPE.id(username),
                                                           USER_EXO_HANDLED_ATTRIBUTE);
    if (settingValue == null || settingValue.getValue() == null) {
      return null;
    }
    long time = Long.parseLong(settingValue.getValue().toString());
    return new Date(time);
  }

  /**
   * @param organizationService
   * @param username
   * @param name
   * @param value
   * @throws Exception
   */
  public static void setUserArrtibute(OrganizationService organizationService,
                                      String username,
                                      String name,
                                      String value) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      if (USER_EXCHANGE_PASSWORD_ATTRIBUTE.equals(name)) {
        value = encodePassword(value);
      }
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      if (userProfile == null) {
        userProfile = organizationService.getUserProfileHandler().createUserProfileInstance(username);
        organizationService.getUserProfileHandler().saveUserProfile(userProfile, true);
      }
      userProfile.setAttribute(name, value);
      organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  /**
   * @param organizationService
   * @param username
   * @param name
   * @return
   * @throws Exception
   */
  public static String getUserArrtibute(OrganizationService organizationService, String username, String name) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      String value = null;
      if (userProfile != null) {
        value = userProfile.getAttribute(name);
        if (value != null && USER_EXCHANGE_PASSWORD_ATTRIBUTE.equals(name)) {
          value = decodePassword(value);
        }
      }
      return value;
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  public ExchangeService getService() {
    return service;
  }

  private void synchronizeExoEventsByModificationDate(FolderId folderId,
                                                      List<String> updatedExoEventIDs,
                                                      Date exoLastSyncDate) throws Exception {
    List<CalendarEvent> modifiedCalendarEvents = searchCalendarEventsModifiedSince(getUserCalendarByExchangeFolderId(folderId),
                                                                                   exoLastSyncDate);
    if (LOG.isTraceEnabled() && modifiedCalendarEvents.size() > 0) {
      LOG.trace("Check exo user calendar for user '{}' since '{}', items found: {}",
                username,
                exoLastSyncDate,
                modifiedCalendarEvents.size());
    }
    for (CalendarEvent calendarEvent : modifiedCalendarEvents) {
      // If modified with synchronization, ignore
      if (updatedExoEventIDs.contains(calendarEvent.getId())) {
        continue;
      }
      String exoMasterId = null;
      if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
        exoMasterId = exoStorageService.getExoEventMasterRecurenceByOriginalUUID(calendarEvent.getOriginalReference());
        if (exoMasterId == null) {
          LOG.error("No master Id was found for occurence: " + calendarEvent.getSummary() + " with recurrenceId = "
              + calendarEvent.getRecurrenceId() + ". The event will not be updated.");
        }
      }
      boolean deleteEvent = exchangeStorageService.updateOrCreateExchangeAppointment(username,
                                                                                     service,
                                                                                     calendarEvent,
                                                                                     exoMasterId,
                                                                                     this::appointmentUpdated);
      if (deleteEvent) {
        exoStorageService.deleteEvent(username, calendarEvent);
      }
      updatedExoEventIDs.add(calendarEvent.getId());
    }
  }

  private FindItemsResults<Item> getItems(FolderId parentFolderId, ItemView view) throws Exception {
    try {
      return service.findItems(parentFolderId, view);
    } catch (Exception e) {
      LOG.warn("Error while paging results: page = " + view.getOffset() + " page size = " + view.getPageSize(), e);
    }
    return null;
  }

  private List<CalendarEvent> searchCalendarEventsModifiedSince(Calendar calendar, Date date) throws Exception {
    if (date == null) {
      return exoStorageService.getAllExoEvents(username, calendar);
    }
    return exoStorageService.findExoEventsModifiedSince(username, calendar, date);
  }

  private List<Item> searchAllAppointmentsModifiedSince(FolderId parentFolderId, Date date) throws Exception {
    if (date == null) {
      throw new IllegalArgumentException("since date is null");
    }

    java.util.Calendar calendar = java.util.Calendar.getInstance();
    calendar.setTime(date);

    ItemView view = new ItemView(100);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId,
                                                           new SearchFilter.IsGreaterThan(ItemSchema.LastModifiedTime,
                                                                                          calendar.getTime()),
                                                           view);
    int totalCount = findResults.getTotalCount();
    if (LOG.isTraceEnabled() && totalCount > 0) {
      LOG.trace("Check exchange user calendar for user '{}' since '{}', items found: {}", username, date, totalCount);
    }
    return findResults.getItems();
  }

  private static void initCodec() throws TokenServiceInitializationException {
    String builderType = PropertyManager.getProperty("gatein.codec.builderclass");
    Map<String, String> config = new HashMap<>();

    if (builderType != null) {
      // If there is config for codec in configuration.properties, we read the
      // config parameters from config file
      // referenced in configuration.properties
      String configFilePath = PropertyManager.getProperty("gatein.codec.config");
      File configFile = new File(configFilePath);
      try (InputStream in = new FileInputStream(configFile)) {
        Properties properties = new Properties();
        properties.load(in);
        for (Map.Entry<?, ?> entry : properties.entrySet()) {
          config.put((String) entry.getKey(), (String) entry.getValue());
        }
        config.put("gatein.codec.config.basedir", configFile.getParentFile().getAbsolutePath());
      } catch (IOException e) {
        throw new TokenServiceInitializationException("Failed to read the config parameters from file '" + configFilePath + "'.",
                                                      e);
      }
    } else {
      // If there is no config for codec in configuration.properties, we
      // generate key if it does not exist and setup the
      // default config
      builderType = "org.exoplatform.web.security.codec.JCASymmetricCodecBuilder";
      String gtnConfDir = PropertyManager.getProperty("gatein.conf.dir");
      if (gtnConfDir == null || gtnConfDir.length() == 0) {
        throw new TokenServiceInitializationException("'gatein.conf.dir' property must be set.");
      }
      File f = new File(gtnConfDir + "/codec/codeckey.txt");
      if (!f.exists()) {
        File codecDir = f.getParentFile();
        if (!codecDir.exists()) {
          codecDir.mkdir();
        }
        try (OutputStream out = new FileOutputStream(f)) {
          KeyGenerator keyGen = KeyGenerator.getInstance("AES");
          keyGen.init(128);
          SecretKey key = keyGen.generateKey();
          KeyStore store = KeyStore.getInstance("JCEKS");
          store.load(null, GTN_STORE_PASS.toCharArray());
          store.setEntry(GTN_KEY, new KeyStore.SecretKeyEntry(key), new KeyStore.PasswordProtection("gtnKeyPass".toCharArray()));
          store.store(out, GTN_STORE_PASS.toCharArray());
        } catch (Exception e) {
          throw new TokenServiceInitializationException(e);
        }
      }
      config.put("gatein.codec.jca.symmetric.keyalg", "AES");
      config.put("gatein.codec.jca.symmetric.keystore", "codeckey.txt");
      config.put("gatein.codec.jca.symmetric.storetype", "JCEKS");
      config.put("gatein.codec.jca.symmetric.alias", GTN_KEY);
      config.put("gatein.codec.jca.symmetric.keypass", "gtnKeyPass");
      config.put("gatein.codec.jca.symmetric.storepass", GTN_STORE_PASS);
      config.put("gatein.codec.config.basedir", f.getParentFile().getAbsolutePath());
    }

    try {
      codec = Class.forName(builderType).asSubclass(AbstractCodecBuilder.class).newInstance().build(config);
    } catch (Exception e) {
      throw new TokenServiceInitializationException("Could not initialize CookieTokenService.codec.", e);
    }
  }

  private static String decodePassword(String password) throws TokenServiceInitializationException {
    if (codec == null) {
      initCodec();
    }
    return codec.decode(password);
  }

  private static String encodePassword(String password) throws TokenServiceInitializationException {
    if (codec == null) {
      initCodec();
    }
    return codec.encode(password);
  }

  /**
   * Make sure that eXo Event has the same modification date than Exchange event
   * 
   * @param appointment
   * @return
   */
  public Boolean appointmentUpdated(Appointment appointment) {
    try {
      // Refresh Appointment
      appointment = (Appointment) exchangeStorageService.getItem(service, appointment.getId());

      CalendarEvent event = exoStorageService.getEventByAppointmentId(username, appointment.getId().getUniqueId());
      if (event != null) {
        exoStorageService.updateModifiedDateOfEvent(username, event, appointment.getLastModifiedTime());
      }
    } catch (Exception e) {
      LOG.warn("Error occurred while updating eXo Event last updated date", e);
      return false;
    }
    return true;
  }
}

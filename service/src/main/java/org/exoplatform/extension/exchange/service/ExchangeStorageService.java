package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.TimeZone;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.AppointmentSchema;
import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.CalendarFolder;
import microsoft.exchange.webservices.data.ConflictResolutionMode;
import microsoft.exchange.webservices.data.DeleteMode;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FindFoldersResults;
import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.FolderView;
import microsoft.exchange.webservices.data.Item;
import microsoft.exchange.webservices.data.ItemId;
import microsoft.exchange.webservices.data.PropertySet;
import microsoft.exchange.webservices.data.ServiceResponseException;
import microsoft.exchange.webservices.data.TimeZoneDefinition;
import microsoft.exchange.webservices.data.WellKnownFolderName;

import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class ExchangeStorageService implements Serializable {
  private static final long serialVersionUID = 6348129698208975430L;

  private final static Log LOG = ExoLogger.getLogger(ExchangeStorageService.class);

  private OrganizationService organizationService;
  private CorrespondenceService correspondenceService;

  public ExchangeStorageService(OrganizationService organizationService, CorrespondenceService correspondenceService) {
    this.organizationService = organizationService;
    this.correspondenceService = correspondenceService;
  }

  /**
   * 
   * Gets list of personnal Exchange Calendars.
   * 
   * @return list of FolderId
   * @throws Exception
   */
  public List<FolderId> getAllExchangeCalendars(ExchangeService service) throws Exception {
    List<FolderId> calendarFolderIds = new ArrayList<FolderId>();
    CalendarFolder calendarRootFolder = CalendarFolder.bind(service, WellKnownFolderName.Calendar);

    calendarFolderIds.add(calendarRootFolder.getId());
    List<Folder> calendarfolders = searchSubFolders(service, calendarRootFolder.getId());

    if (calendarfolders != null && !calendarfolders.isEmpty()) {
      for (Folder tmpFolder : calendarfolders) {
        calendarFolderIds.add(tmpFolder.getId());
      }
    }
    return calendarFolderIds;
  }

  /**
   * 
   * @param username
   * @param service
   * @param event
   * @param exoMasterId
   * @param userCalendarTimeZone
   * @return true if the CalenarEvent have to be deleted
   * @throws Exception
   */
  public boolean updateOrCreateExchangeAppointment(String username, ExchangeService service, CalendarEvent event, String exoMasterId, TimeZone userCalendarTimeZone,
      List<CalendarEvent> eventsToUpdateModifiedTime) throws Exception {
    if (event == null) {
      return false;
    }
    String folderIdString = correspondenceService.getCorrespondingId(username, event.getCalendarId());
    if (folderIdString == null || folderIdString.isEmpty()) {
      LOG.trace("eXo Calendar with id '" + event.getCalendarId() + "' is not synhronized with Exchange, ignore Event:" + event.getSummary());
      return false;
    }

    String itemId = correspondenceService.getCorrespondingId(username, event.getId());
    boolean isNew = true;
    Appointment appointment = null;
    if (itemId != null) {
      try {
        appointment = Appointment.bind(service, ItemId.getItemIdFromString(itemId));
        isNew = false;
      } catch (ServiceResponseException e) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Item was not bound, it was deleted or not yet created:" + event.getId());
        }
        correspondenceService.deleteCorrespondingId(username, event.getId());
      }
    }

    if (event.getRecurrenceId() == null && (event.getRepeatType() == null || event.getRepeatType().equals(CalendarEvent.RP_NOREPEAT))) {
      if (isNew) {
        // Checks if this event was already in Exchange, if it's the case, it
        // means that the item was not found because the user has removed it
        // from
        // Exchange
        if (CalendarConverterService.isExchangeEventId(event.getId())) {
          LOG.error("Conflict in modification, inconsistant data, the event was deleted in Exchange but seems always in eXo, the event will be deleted from Exchange.");
          deleteAppointmentByExoEventId(username, service, event.getId(), event.getCalendarId());
          return false;
        }
        appointment = new Appointment(service);
      }
      CalendarConverterService
          .convertExoToExchangeEvent(appointment, event, username, organizationService.getUserHandler(), getTimeZoneDefinition(service, userCalendarTimeZone), userCalendarTimeZone);
    } else {
      if ((event.getRecurrenceId() != null && !event.getRecurrenceId().isEmpty()) || (event.getIsExceptionOccurrence() != null && event.getIsExceptionOccurrence())) {
        if (isNew) {
          String exchangeMasterId = correspondenceService.getCorrespondingId(username, exoMasterId);
          Appointment tmpAppointment = getAppointmentOccurence(service, exchangeMasterId, event.getRecurrenceId());
          if (tmpAppointment != null) {
            appointment = tmpAppointment;
            isNew = false;
          } else {
            appointment = new Appointment(service);
          }
        }
        CalendarConverterService.convertExoToExchangeOccurenceEvent(appointment, event, username, organizationService.getUserHandler(), getTimeZoneDefinition(service, userCalendarTimeZone),
            userCalendarTimeZone);
      } else {
        if (isNew) {
          // Checks if this event was already in Exchange, if it's the case, it
          // means that the item was not found because the user has removed it
          // from Exchange
          if (CalendarConverterService.isExchangeEventId(event.getId())) {
            LOG.error("Conflict in modification, inconsistant data, the event was deleted in Exchange but seems always in eXo, the event will be deleted from Exchange.");
            deleteAppointmentByExoEventId(username, service, event.getId(), event.getCalendarId());
            return false;
          }
          appointment = new Appointment(service);
        }
        List<Appointment> toDeleteOccurences = CalendarConverterService.convertExoToExchangeMasterRecurringCalendarEvent(appointment, event, username, organizationService.getUserHandler(),
            getTimeZoneDefinition(service, userCalendarTimeZone), userCalendarTimeZone);

        if (toDeleteOccurences != null && !toDeleteOccurences.isEmpty()) {
          for (Appointment occAppointment : toDeleteOccurences) {
            // Verify if deleted occurences is an exception existing occurence
            // or not
            String exoId = correspondenceService.getCorrespondingId(username, occAppointment.getId().getUniqueId());
            if (exoId == null) {
              deleteAppointment(username, service, occAppointment.getId());
            }
          }
        }
      }
    }
    if (isNew) {
      LOG.info("Create Exchange Appointment: " + event.getSummary());
      FolderId folderId = FolderId.getFolderIdFromString(folderIdString);
      appointment.save(folderId);
      correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
    } else /*
            * test if appointment wasn't deleted by previous
            * 'toDeleteOccurences' List
            */if (correspondenceService.getCorrespondingId(username, event.getId()) != null) {
      LOG.info("Update Exchange Appointment: " + event.getSummary());
      appointment.update(ConflictResolutionMode.AlwaysOverwrite);
      correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
    }
    if (eventsToUpdateModifiedTime != null) {
      eventsToUpdateModifiedTime.add(event);
    }
    return false;
  }

  /**
   * 
   * @param username
   * @param service
   * @param eventId
   * @param calendarId
   * @throws Exception
   */
  public void deleteAppointmentByExoEventId(String username, ExchangeService service, String eventId, String calendarId) throws Exception {
    String itemId = correspondenceService.getCorrespondingId(username, eventId);
    if (itemId == null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("The event was deleted from eXo but seems don't have corresponding Event in Exchange, ignore.");
      }
    } else {
      // Verify that calendar is synchronized
      if (correspondenceService.getCorrespondingId(username, calendarId) == null) {
        LOG.warn("Calendar with id '" + calendarId + "' seems not synchronized with exchange.");

        // delete correspondance for safe operations while sync is activated
        correspondenceService.deleteCorrespondingId(username, eventId);
      } else {
        deleteAppointment(username, service, itemId);
      }
    }
  }

  /**
   * 
   * @param username
   * @param service
   * @param itemId
   * @throws Exception
   */
  public void deleteAppointment(String username, ExchangeService service, String itemId) throws Exception {
    deleteAppointment(username, service, ItemId.getItemIdFromString(itemId));
  }

  /**
   * 
   * @param username
   * @param service
   * @param itemId
   * @throws Exception
   */
  public void deleteAppointment(String username, ExchangeService service, ItemId itemId) throws Exception {
    Appointment appointment = null;
    try {
      appointment = Appointment.bind(service, itemId);
      LOG.info("Delete Exchange appointment: " + appointment.getSubject());
      appointment.delete(DeleteMode.HardDelete);
    } catch (ServiceResponseException e) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Exchange Item was not bound, it was deleted or not yet created:" + itemId);
      }
    }
    correspondenceService.deleteCorrespondingId(username, itemId.getUniqueId());
  }

  /**
   * 
   * @param username
   * @param service
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeFolderByCalenarId(String username, ExchangeService service, String calendarId) throws Exception {
    if (CalendarConverterService.isExchangeCalendarId(calendarId)) {
      LOG.warn("Can't delete Exchange Calendar, because it was created on Exchange: " + calendarId);
      return;
    }
    String folderId = correspondenceService.getCorrespondingId(username, calendarId);
    if (folderId == null) {
      LOG.warn("Conflict in modification, inconsistant data, the Calendar was deleted from eXo but seems don't have corresponding Folder in Exchange, ignore.");
      return;
    } else {
      Folder folder = null;
      try {
        folder = Folder.bind(service, FolderId.getFolderIdFromString(folderId));
        LOG.trace("Delete Exchange folder: " + folder.getDisplayName());
        folder.delete(DeleteMode.MoveToDeletedItems);
      } catch (ServiceResponseException e) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Exchange Folder was not bound, it was deleted or not yet created:" + folderId);
        }
      }
      correspondenceService.deleteCorrespondingId(username, calendarId);
    }
  }

  private TimeZoneDefinition getTimeZoneDefinition(ExchangeService service, TimeZone userCalendarTimeZone) {
    TimeZoneDefinition serverTimeZoneDefinition = null;
    Iterator<TimeZoneDefinition> timeZoneDefinitions = service.getServerTimeZones().iterator();
    while (timeZoneDefinitions.hasNext()) {
      TimeZoneDefinition timeZoneDefinition = (TimeZoneDefinition) timeZoneDefinitions.next();
      if (timeZoneDefinition.getId().equals(userCalendarTimeZone.getID())) {
        serverTimeZoneDefinition = timeZoneDefinition;
        break;
      }
    }
    return serverTimeZoneDefinition;
  }

  private List<Folder> searchSubFolders(ExchangeService service, FolderId parentFolderId) throws Exception {
    FolderView view = new FolderView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindFoldersResults findResults = service.findFolders(parentFolderId, view);
    return findResults.getFolders();
  }

  private Appointment getAppointmentOccurence(ExchangeService service, String exchangeMasterId, String recurrenceId) throws Exception {
    Appointment masterAppointment = Appointment.bind(service, ItemId.getItemIdFromString(exchangeMasterId), new PropertySet(AppointmentSchema.Recurrence));
    return CalendarConverterService.getAppointmentOccurence(masterAppointment, recurrenceId);
  }

  /**
   * 
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(ExchangeService service, String folderId) throws Exception {
    return getExchangeCalendar(service, FolderId.getFolderIdFromString(folderId));
  }

  /**
   * 
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(ExchangeService service, FolderId folderId) throws Exception {
    CalendarFolder folder = null;
    try {
      folder = CalendarFolder.bind(service, folderId);
    } catch (ServiceResponseException e) {
      LOG.warn("Can't get Folder identified by id: " + folderId.getUniqueId());
    }
    return folder;
  }

  /**
   * 
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public Appointment getAppointment(ExchangeService service, String appointmentId) throws Exception {
    return getAppointment(service, ItemId.getItemIdFromString(appointmentId));
  }

  /**
   * 
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public Appointment getAppointment(ExchangeService service, ItemId appointmentId) throws Exception {
    Appointment appointment = null;
    try {
      appointment = Appointment.bind(service, appointmentId);
    } catch (ServiceResponseException e) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Can't get appointment identified by id: " + appointmentId.getUniqueId());
      }
    }
    return appointment;
  }

  /**
   * 
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public Item getItem(ExchangeService service, String itemId) throws Exception {
    return getItem(service, ItemId.getItemIdFromString(itemId));
  }

  /**
   * 
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public Item getItem(ExchangeService service, ItemId itemId) throws Exception {
    Item item = null;
    try {
      item = Item.bind(service, itemId);
    } catch (ServiceResponseException e) {
      LOG.warn("Can't get item identified by id: " + itemId.getUniqueId());
    }
    return item;
  }
}
package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.List;
import java.util.TimeZone;

import javax.jcr.Node;
import javax.jcr.NodeIterator;
import javax.jcr.query.Query;
import javax.jcr.query.QueryManager;
import javax.jcr.query.QueryResult;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.AppointmentSchema;
import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.PropertySet;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.Utils;
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.calendar.service.impl.JCRDataStorage;
import org.exoplatform.commons.utils.ISO8601;
import org.exoplatform.extension.exchange.listener.CalendarCreateUpdateAction;
import org.exoplatform.extension.exchange.service.util.CalendarConverterService;
import org.exoplatform.services.jcr.ext.common.SessionProvider;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class ExoStorageService implements Serializable {
  private static final long serialVersionUID = 6614108102985034995L;

  private final static Log LOG = ExoLogger.getLogger(ExoStorageService.class);
  private final static DateFormat EXCLUDE_ID_FORMAT_FIRST_CHARS = new SimpleDateFormat("yyyyMMdd");

  private JCRDataStorage storage;
  private OrganizationService organizationService;
  private CorrespondenceService correspondenceService;

  public ExoStorageService(OrganizationService organizationService, CalendarService calendarService, CorrespondenceService correspondenceService) {
    this.storage = ((CalendarServiceImpl) calendarService).getDataStorage();
    this.organizationService = organizationService;
    this.correspondenceService = correspondenceService;
  }

  /**
   * 
   * Deletes eXo Calendar Event that corresponds to given appointment Id.
   * 
   * @param appointmentId
   * @param username
   * @throws Exception
   */
  public void deleteEventByAppointmentID(String appointmentId, String username) throws Exception {
    CalendarEvent calendarEvent = getEventByAppointmentId(username, appointmentId);
    if (calendarEvent != null) {
      deleteEvent(username, calendarEvent);
    }
  }

  /**
   * 
   * Deletes eXo Calendar Event.
   * 
   * @param username
   * @param calendarEvent
   * @throws Exception
   */
  public void deleteEvent(String username, CalendarEvent calendarEvent) throws Exception {
    if (calendarEvent == null) {
      LOG.warn("Event is null, can't delete it for username: " + username);
      return;
    }

    if ((calendarEvent.getRepeatType() == null || calendarEvent.getRepeatType().equals(CalendarEvent.RP_NOREPEAT))
        && (calendarEvent.getIsExceptionOccurrence() == null || !calendarEvent.getIsExceptionOccurrence())) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Delete user calendar event: " + calendarEvent.getSummary());
      }
      storage.removeUserEvent(username, calendarEvent.getCalendarId(), calendarEvent.getId());
      // Remove correspondence between exo and exchange IDs
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    } else if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Delete user calendar event exceptional occurence: " + calendarEvent.getSummary() + ", id=" + calendarEvent.getRecurrenceId());
      }
      storage.removeUserEvent(username, calendarEvent.getCalendarId(), calendarEvent.getId());
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    } else if (calendarEvent.getRecurrenceId() != null && !calendarEvent.getRecurrenceId().isEmpty()) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Delete user calendar event occurence from series: " + calendarEvent.getSummary() + " with id : " + calendarEvent.getRecurrenceId());
      }
      storage.removeOccurrenceInstance(username, calendarEvent);
    } else {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Delete user calendar event series: " + calendarEvent.getSummary());
      }
      storage.removeRecurrenceSeries(username, calendarEvent);
      // Remove correspondence between exo and exchange IDs
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    }
  }

  /**
   * 
   * Delete eXo Calendar.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  public boolean deleteCalendar(String username, String folderId) throws Exception {
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    if (calendarId == null) {
      calendarId = CalendarConverterService.getCalendarId(folderId);
    }
    List<CalendarEvent> events = getUserCalendarEvents(username, folderId);
    if (events == null) {
      return false;
    }
    for (CalendarEvent calendarEvent : events) {
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    }
    storage.removeUserCalendar(username, calendarId);
    correspondenceService.deleteCorrespondingId(username, folderId, calendarId);
    return true;
  }

  /**
   * 
   * Gets User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  public Calendar getUserCalendar(String username, String folderId) throws Exception {
    return getUserCalendar(username, folderId, true);
  }

  /**
   * 
   * Gets User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @param deleteIfCorrespondentExists
   * @return
   * @throws Exception
   */
  public Calendar getUserCalendar(String username, String folderId, boolean deleteIfCorrespondentExists) throws Exception {
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    Calendar calendar = null;
    if (calendarId != null) {
      calendar = storage.getUserCalendar(username, calendarId);
      if (calendar == null && deleteIfCorrespondentExists) {
        correspondenceService.deleteCorrespondingId(username, folderId);
      }
    }
    return calendar;
  }

  /**
   * 
   * Gets User Calendar identified by Exchange folder Id, or creates it if not
   * existing.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  public Calendar getOrCreateUserCalendar(String username, Folder folder) throws Exception {
    Calendar calendar = getUserCalendar(username, folder.getId().getUniqueId(), false);
    String calendarId = CalendarConverterService.getCalendarId(folder.getId().getUniqueId());
    if (calendar == null) {
      Calendar tmpCalendar = storage.getUserCalendar(username, calendarId);
      if (tmpCalendar != null) {
        // Renew Calendar
        storage.removeUserCalendar(username, calendarId);
      }
    }
    if (calendar == null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Create user calendar from Exchange: " + folder.getDisplayName());
      }

      calendar = new Calendar();
      calendar.setId(calendarId);
      calendar.setName(CalendarConverterService.getCalendarName(folder.getDisplayName()));
      calendar.setCalendarOwner(username);
      calendar.setDataInit(false);
      calendar.setEditPermission(new String[] { "any read" });
      calendar.setCalendarColor(Calendar.COLORS[(int) (Math.random() * Calendar.COLORS.length)]);

      storage.saveUserCalendar(username, calendar, true);

      // Set IDs correspondence
      correspondenceService.setCorrespondingId(username, calendar.getId(), folder.getId().getUniqueId());
    }
    return calendar;
  }

  /**
   * 
   * Gets Events from User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> getUserCalendarEvents(String username, String folderId) throws Exception {
    List<CalendarEvent> userEvents = null;
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    if (calendarId == null) {
      calendarId = CalendarConverterService.getCalendarId(folderId);
    }
    Calendar calendar = storage.getUserCalendar(username, calendarId);
    if (calendar != null) {
      List<String> calendarIds = new ArrayList<String>();
      calendarIds.add(calendarId);
      userEvents = storage.getUserEventByCalendar(username, calendarIds);
    }
    return userEvents;
  }

  /**
   * 
   * Updates existing eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @param timeZone
   * @throws Exception
   */
  public List<CalendarEvent> updateEvent(Appointment appointment, String username, TimeZone timeZone) throws Exception {
    return createOrUpdateEvent(appointment, username, false, timeZone);
  }

  /**
   * 
   * Create non existing eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @param timeZone
   * @throws Exception
   */
  public List<CalendarEvent> createEvent(Appointment appointment, String username, TimeZone timeZone) throws Exception {
    return createOrUpdateEvent(appointment, username, true, timeZone);
  }

  /**
   * 
   * Creates or updates eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @param timeZone
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> createOrUpdateEvent(Appointment appointment, String username, TimeZone timeZone) throws Exception {
    boolean isNew = correspondenceService.getCorrespondingId(username, appointment.getId().getUniqueId()) == null;
    if (!isNew) {
      CalendarEvent event = getEventByAppointmentId(username, appointment.getId().getUniqueId());
      if (event == null) {
        isNew = true;
        correspondenceService.deleteCorrespondingId(username, appointment.getId().getUniqueId());
      }
    }
    return createOrUpdateEvent(appointment, username, isNew, timeZone);
  }

  /**
   * 
   * @param username
   * @param appointmentId
   * @return
   * @throws Exception
   */
  public CalendarEvent getEventByAppointmentId(String username, String appointmentId) throws Exception {
    String calEventId = correspondenceService.getCorrespondingId(username, appointmentId);
    CalendarEvent event = storage.getEvent(username, calEventId);
    if (event == null && calEventId != null) {
      correspondenceService.deleteCorrespondingId(username, appointmentId);
    }
    return event;
  }

  /**
   * 
   * @param eventNode
   * @return
   * @throws Exception
   */
  public CalendarEvent getExoEventByNode(Node eventNode) throws Exception {
    return storage.getEvent(eventNode);
  }

  /**
   * 
   * @param uuid
   * @return
   * @throws Exception
   */
  public String getExoEventMasterRecurenceByOriginalUUID(String uuid) throws Exception {
    Node node = storage.getSession(SessionProvider.createSystemProvider()).getNodeByUUID(uuid);
    if (node == null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("No original recurrent node was found with UUID: " + uuid);
      }
      return null;
    } else {
      return node.getName();
    }
  }

  /**
   * 
   * @param username
   * @param calendar
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> getAllExoEvents(String username, Calendar calendar) throws Exception {
    List<String> calendarIds = Collections.singletonList(calendar.getId());
    return storage.getUserEventByCalendar(username, calendarIds);
  }

  /**
   * 
   * @param username
   * @param calendar
   * @param date
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> findExoEventsModifiedSince(String username, Calendar calendar, Date date) throws Exception {
    Node calendarHome = storage.getUserCalendarHome(username);
    if (calendarHome.hasNode(calendar.getId())) {
      calendarHome = calendarHome.getNode(calendar.getId());
    }
    java.util.Calendar dateCalendar = java.util.Calendar.getInstance();
    dateCalendar.setTime(date);
    return getEventsByType(calendarHome, Calendar.TYPE_PRIVATE, dateCalendar);
  }

  private List<CalendarEvent> getEventsByType(Node calendarHome, int type, java.util.Calendar date) throws Exception {
    List<CalendarEvent> events = new ArrayList<CalendarEvent>();
    QueryManager qm = calendarHome.getSession().getWorkspace().getQueryManager();
    Query query = qm.createQuery("select * from exo:calendarEvent where (jcr:path like '" + calendarHome.getPath() + "/%') and (exo:lastModifiedDate > TIMESTAMP '" + ISO8601.format(date)
        + "')", Query.SQL);
    QueryResult result = query.execute();
    NodeIterator it = result.getNodes();
    CalendarEvent calEvent;
    while (it.hasNext()) {
      calEvent = storage.getEvent(it.nextNode());
      calEvent.setCalType(String.valueOf(type));
      events.add(calEvent);
    }
    return events;
  }

  public void updateModifiedDateOfEvent(String username, CalendarEvent event) throws Exception {
    Node node = storage.getCalendarEventNode(username, event.getCalType(), event.getCalendarId(), event.getId());
    modifyUpdateDate(node);
    if (event.getOriginalReference() != null && !event.getOriginalReference().isEmpty()) {
      Node masterNode = storage.getSession(SessionProvider.createSystemProvider()).getNodeByUUID(event.getOriginalReference());
      modifyUpdateDate(masterNode);
    }
  }

  private void modifyUpdateDate(Node node) throws Exception {
    if (!node.isNodeType("exo:datetime")) {
      if (node.canAddMixin("exo:datetime")) {
        node.addMixin("exo:datetime");
      }
      node.setProperty("exo:dateCreated", new GregorianCalendar());
    }
    node.setProperty("exo:dateModified", new GregorianCalendar());
    node.save();
  }

  private List<CalendarEvent> createOrUpdateEvent(Appointment appointment, String username, boolean isNew, TimeZone timeZone) throws Exception {
    Calendar calendar = getUserCalendar(username, appointment.getParentFolderId().getUniqueId());
    if (calendar == null) {
      LOG.warn("Attempting to synchronize an event without existing associated eXo Calendar.");
      return null;
    }
    List<CalendarEvent> updatedEvents = new ArrayList<CalendarEvent>();

    if (appointment.getAppointmentType() != null) {
      switch (appointment.getAppointmentType()) {
      case Single: {
        CalendarEvent event = null;
        if (isNew) {
          event = new CalendarEvent();
          event.setCalendarId(calendar.getId());
          updatedEvents.add(event);
        } else {
          event = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          updatedEvents.add(event);
          if (CalendarConverterService.verifyModifiedDatesConflict(event, appointment)) {
            if (LOG.isTraceEnabled()) {
              LOG.trace("Attempting to update eXo Event with Exchange Event, but modification date of eXo is after, ignore updating.");
            }
            return updatedEvents;
          }
        }

        if (LOG.isTraceEnabled()) {
          if (isNew) {
            LOG.trace("Create user calendar event: " + appointment.getSubject());
          } else {
            LOG.trace("Update user calendar event: " + appointment.getSubject());
          }
        }

        CalendarConverterService.convertExchangeToExoEvent(event, appointment, username, storage, organizationService.getUserHandler(), timeZone);
        event.setRepeatType(CalendarEvent.RP_NOREPEAT);
        CalendarCreateUpdateAction.IGNORE_UPDATE.set(true);
        try {
          storage.saveUserEvent(username, calendar.getId(), event, isNew);
        } finally {
          CalendarCreateUpdateAction.IGNORE_UPDATE.set(false);
        }
        correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
      }
        break;
      case Exception:
        throw new IllegalStateException("The appointment is an exception occurence of this event >> '" + appointment.getSubject() + "'. start:" + appointment.getStart() + ", end : "
            + appointment.getEnd() + ", occurence: " + appointment.getAppointmentSequenceNumber());
      case RecurringMaster: {
        // Master recurring event
        CalendarEvent masterEvent = null;
        Date orginialStartDate = null;
        if (isNew) {
          masterEvent = new CalendarEvent();
          updatedEvents.add(masterEvent);
        } else {
          masterEvent = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          updatedEvents.add(masterEvent);
          orginialStartDate = masterEvent.getFromDateTime();
        }

        // FIXME there is a bug in Exchange modification time of the server:
        // Adding a recurrent Item + delete last occurence => last modified
        // date isn't updated. So we test here if the last occurence was deleted
        // or not
        //
        // Begin workaround
        boolean isLastOccurenceDeleted = false;
        appointment = Appointment.bind(appointment.getService(), appointment.getId(), new PropertySet(AppointmentSchema.Recurrence));
        if (appointment.getRecurrence().hasEnd()) {
          Date recEndDate = appointment.getRecurrence().getEndDate();

          appointment = Appointment.bind(appointment.getService(), appointment.getId(), new PropertySet(BasePropertySet.FirstClassProperties));
          if (recEndDate == null) {
            LOG.warn("Inconsistent data delivered by MS Exchange. The recurrent Event has end but end date is null: '" + appointment.getSubject() + "', start:" + appointment.getStart() + ", end : "
                + appointment.getEnd());
          } else {
            Appointment tmpAppointment = Appointment.bind(appointment.getService(), appointment.getId(), new PropertySet(AppointmentSchema.LastOccurrence));
            if (tmpAppointment.getLastOccurrence() == null) {
              LOG.warn("Can't find last occurence of recurrent Event : '" + appointment.getSubject() + "', start:" + appointment.getStart() + ", end : " + appointment.getEnd());
            } else {
              isLastOccurenceDeleted = tmpAppointment.getLastOccurrence().getEnd().getTime() < recEndDate.getTime();

              if (isLastOccurenceDeleted && masterEvent.getExcludeId() != null) {
                String pattern = EXCLUDE_ID_FORMAT_FIRST_CHARS.format(recEndDate);
                int i = 0;
                while (isLastOccurenceDeleted && i < masterEvent.getExcludeId().length) {
                  isLastOccurenceDeleted = !masterEvent.getExcludeId()[i].startsWith(pattern);
                  i++;
                }
              }
            }
          }
        }
        appointment = Appointment.bind(appointment.getService(), appointment.getId(), new PropertySet(BasePropertySet.FirstClassProperties));
        // End workaround

        if (!isLastOccurenceDeleted && !isNew && CalendarConverterService.verifyModifiedDatesConflict(masterEvent, appointment)) {
          if (LOG.isTraceEnabled()) {
            LOG.trace("Attempting to update eXo Event with Exchange Event, but modification date of eXo is after, ignore updating.");
          }
          return updatedEvents;
        } else {
          if (LOG.isTraceEnabled()) {
            if (isNew) {
              LOG.trace("Create recurrent user calendar event: " + appointment.getSubject());
            } else {
              LOG.trace("Update recurrent user calendar event: " + appointment.getSubject());
            }
          }

          masterEvent.setCalendarId(calendar.getId());
          CalendarConverterService.convertExchangeToExoMasterRecurringCalendarEvent(masterEvent, appointment, username, storage, organizationService.getUserHandler(), timeZone);
          if (isNew) {
            correspondenceService.setCorrespondingId(username, masterEvent.getId(), appointment.getId().getUniqueId());
          } else if (!CalendarConverterService.isSameDate(orginialStartDate, masterEvent.getFromDateTime())) {
            if (masterEvent.getExcludeId() == null) {
              masterEvent.setExcludeId(new String[0]);
            }
          }

          CalendarCreateUpdateAction.IGNORE_UPDATE.set(true);
          try {
            storage.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
          } finally {
            CalendarCreateUpdateAction.IGNORE_UPDATE.set(false);
          }
        }
        List<CalendarEvent> exceptionalEventsToUpdate = new ArrayList<CalendarEvent>();
        List<String> occAppointmentIDs = new ArrayList<String>();
        // Deleted execptional occurences events.
        List<CalendarEvent> toDeleteEvents = CalendarConverterService.convertExchangeToExoOccurenceEvent(masterEvent, exceptionalEventsToUpdate, occAppointmentIDs, appointment, username, storage,
            organizationService.getUserHandler(), correspondenceService, timeZone);
        if (exceptionalEventsToUpdate != null && !exceptionalEventsToUpdate.isEmpty()) {
          storage.updateOccurrenceEvent(calendar.getId(), calendar.getId(), masterEvent.getCalType(), masterEvent.getCalType(), exceptionalEventsToUpdate, username);

          // Set correspondance IDs
          Iterator<CalendarEvent> eventsIterator = exceptionalEventsToUpdate.iterator();
          Iterator<String> occAppointmentIdIterator = occAppointmentIDs.iterator();
          while (eventsIterator.hasNext()) {
            CalendarEvent calendarEvent = eventsIterator.next();
            String occAppointmentId = occAppointmentIdIterator.next();
            correspondenceService.setCorrespondingId(username, calendarEvent.getId(), occAppointmentId);
          }
          updatedEvents.addAll(exceptionalEventsToUpdate);
        }
        if (toDeleteEvents != null && !toDeleteEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : toDeleteEvents) {
            deleteEvent(username, calendarEvent);
          }
        }
      }
        break;
      case Occurrence:
        LOG.warn("The appointment is an occurence of this event >> '" + appointment.getSubject() + "'. start:" + appointment.getStart() + ", end : " + appointment.getEnd() + ", occurence: "
            + appointment.getAppointmentSequenceNumber());
      }
    }
    return updatedEvents;
  }
}
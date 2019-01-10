package org.exoplatform.extension.exchange.service.util;

import java.net.URI;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.temporal.ChronoUnit;
import java.util.*;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.*;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FolderView;

public class ExchangeStaticDataInjector {

  public static void main(String[] args) {
    try (ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)) {
      connectToExchangeServer(service, args);
      List<FolderId> allExchangeCalendars = getAllExchangeCalendars(service);
      FolderId folderId = allExchangeCalendars.get(0);
      LocalDateTime localDateTime = LocalDateTime.now();
      String eventsCountString = args[4];
      int eventsCount = Integer.parseInt(eventsCountString);
      for (int i = 1; i <= eventsCount; i++) {
        LocalDateTime startDate = localDateTime.minus(i, ChronoUnit.DAYS);
        LocalDateTime endDate = startDate.plus(1, ChronoUnit.HOURS);
        Appointment appointment = new Appointment(service);
        appointment.setStart(new Date(startDate.toEpochSecond(ZoneOffset.UTC) * 1000));
        appointment.setEnd(new Date(endDate.toEpochSecond(ZoneOffset.UTC) * 1000));
        appointment.setSubject("Test " + i + "/" + eventsCount);
        appointment.setSensitivity(Sensitivity.Normal);
        appointment.save(folderId);
      }
    } catch (Exception e) {
      throw new IllegalStateException("Error retrieving exchange informations", e);
    }
  }

  private static void connectToExchangeServer(ExchangeService service, String[] args) {
    String userName = args[0];
    String password = args[1];
    String domain = args[2];
    String url = args[3];

    try {
      service.setTimeout(300000);
      ExchangeCredentials credentials = new WebCredentials(userName, password, domain);
      service.setCredentials(credentials);
      service.setUrl(new URI(url));

      service.getInboxRules();
    } catch (Exception e) {
      throw new IllegalStateException("Error connecting to exchange server", e);
    }
  }

  private static List<FolderId> getAllExchangeCalendars(ExchangeService service) throws Exception {
    List<FolderId> calendarFolderIds = new ArrayList<>();
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

  private static List<Folder> searchSubFolders(ExchangeService service, FolderId parentFolderId) throws Exception {
    FolderView view = new FolderView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindFoldersResults findResults = service.findFolders(parentFolderId, view);
    return findResults.getFolders();
  }

}

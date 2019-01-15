package org.exoplatform.extension.exchange.service;

import java.io.*;
import java.util.*;

import javax.jcr.*;

import org.apache.commons.io.output.ByteArrayOutputStream;

import org.exoplatform.calendar.service.Utils;
import org.exoplatform.commons.utils.CommonsUtils;
import org.exoplatform.extension.exchange.service.util.CalendarConverterUtils;
import org.exoplatform.services.jcr.ext.app.SessionProviderService;
import org.exoplatform.services.jcr.ext.common.SessionProvider;
import org.exoplatform.services.jcr.ext.hierarchy.NodeHierarchyCreator;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

public class CorrespondenceService implements Serializable {
  private static final long                serialVersionUID   = 4155183714826625091L;

  private static final Log                 LOG                = ExoLogger.getLogger(CorrespondenceService.class);

  private static final String              EXCHANGE_NODE_NAME = "calendar-exchange-extension";

  // Map of userId, correspondence exchange and eXo Ids
  private Map<String, Properties>          propertiesMap      = new HashMap<>();

  private transient NodeHierarchyCreator   hierarchyCreator;

  private transient SessionProviderService sessionProviderService;

  /**
   * Gets Id of exchange from eXo Calendar or Event Id and vice versa
   * 
   * @param username
   * @param id
   * @return Id of the corresponding element
   * @throws Exception
   */
  public String getCorrespondingId(String username, String id) throws Exception {
    Properties properties = loadCorrespondenceProperties(username);
    return properties.getProperty(id);
  }

  /**
   * Sets Correspondence between IDs
   * 
   * @param username
   * @param exoId
   * @param exchangeId
   * @throws Exception
   */
  public void setCorrespondingId(String username, String exoId, String exchangeId) throws Exception {
    String oldExoId = getCorrespondingId(username, exchangeId);
    String oldExchangeId = getCorrespondingId(username, exoId);
    if ((oldExoId != null && !oldExoId.equals(exoId)) || (oldExchangeId != null && !oldExchangeId.equals(exchangeId))) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Exchange integration, correspondence service : An old existing ID will be replaced by another one.");
      }
      // Make sure no duplicated entry
      deleteCorrespondingId(username, exchangeId, exoId);
    }

    Properties properties = loadCorrespondenceProperties(username);
    properties.setProperty(exchangeId, exoId);
    properties.setProperty(exoId, exchangeId);
    saveProperties(username, properties);
  }

  /**
   * delete Correspondence between IDs
   * 
   * @param username
   * @param exchangeId
   * @param exoId
   * @throws Exception
   */
  public void deleteCorrespondingId(String username, String exchangeId, String exoId) throws Exception {
    Properties properties = loadCorrespondenceProperties(username);
    properties.remove(exchangeId);
    properties.remove(exoId);
    saveProperties(username, properties);
  }

  public void deleteCorrespondingId(String username, String id) throws Exception {
    Properties properties = loadCorrespondenceProperties(username);
    String secondId = properties.getProperty(id);
    if (secondId != null) {
      properties.remove(id);
      properties.remove(secondId);
      saveProperties(username, properties);
    }
  }

  public List<String> getSynchronizedExchangeFolderIds(String username) throws Exception {
    Properties properties = loadCorrespondenceProperties(username);
    List<String> folderIds = new ArrayList<>();
    @SuppressWarnings("unchecked")
    Enumeration<String> enumeration = (Enumeration<String>) properties.propertyNames();
    while (enumeration.hasMoreElements()) {
      String name = enumeration.nextElement();
      if (CalendarConverterUtils.isExchangeCalendarId(name)) {
        folderIds.add(properties.getProperty(name));
      }
    }
    return folderIds;
  }

  private void saveProperties(String username, Properties properties) throws Exception {
    try {
      ByteArrayOutputStream out = new ByteArrayOutputStream();
      properties.store(out, "");
      SessionProvider sessionProvider = getSessionProviderService().getSystemSessionProvider(null);
      Node node = getHierarchyCreator().getUserApplicationNode(sessionProvider, username);
      if (node == null) {
        throw new IllegalStateException("User application node not found. Please fix this and try later.");
      }
      Session session = node.getSession();
      if (!node.hasNode(EXCHANGE_NODE_NAME)) {
        node = node.addNode(EXCHANGE_NODE_NAME, Utils.NT_RESOURCE);
        node.setProperty(Utils.JCR_LASTMODIFIED, java.util.Calendar.getInstance().getTimeInMillis());
        node.setProperty(Utils.JCR_MIMETYPE, "text/plain");
      } else {
        node = node.getNode(EXCHANGE_NODE_NAME);
      }
      node.setProperty(Utils.JCR_DATA, new ByteArrayInputStream(out.toByteArray()));
      session.save();
    } catch (ItemExistsException e) {
      LOG.debug("Error whale saving properties, reattempting", e);
      saveProperties(username, properties);
    }
  }

  private Properties loadCorrespondenceProperties(String username) throws Exception {
    Properties properties = propertiesMap.get(username);
    if (properties == null) {
      properties = new Properties();

      // Load properties from JCR
      SessionProvider sessionProvider = getSessionProviderService().getSystemSessionProvider(null);
      Node node = getHierarchyCreator().getUserApplicationNode(sessionProvider, username);
      if (node == null) {
        throw new IllegalStateException("User application node not found. Please fix this and try later.");
      }
      if (node.hasNode(EXCHANGE_NODE_NAME)) {
        node = node.getNode(EXCHANGE_NODE_NAME);
        InputStream inputStream = node.getProperty(Utils.JCR_DATA).getStream();
        properties.load(inputStream);
      }

      propertiesMap.put(username, properties);
    }
    return properties;
  }

  public NodeHierarchyCreator getHierarchyCreator() {
    if (hierarchyCreator == null) {
      hierarchyCreator = CommonsUtils.getService(NodeHierarchyCreator.class);
    }
    return hierarchyCreator;
  }

  public SessionProviderService getSessionProviderService() {
    if (sessionProviderService == null) {
      sessionProviderService = CommonsUtils.getService(SessionProviderService.class);
    }
    return sessionProviderService;
  }
}

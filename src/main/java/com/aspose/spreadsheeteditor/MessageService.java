package com.aspose.spreadsheeteditor;

import javax.enterprise.context.ApplicationScoped;
import javax.faces.application.FacesMessage;
import javax.faces.context.FacesContext;
import javax.inject.Named;
import org.primefaces.context.RequestContext;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "msg")
@ApplicationScoped
public class MessageService {

    public void sendMessage(String summary, String details) {
        FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(summary, details));
    }

    public void sendMessageDialog(String summary, String details) {
        RequestContext.getCurrentInstance().showMessageInDialog(new FacesMessage(summary, details));
    }
}

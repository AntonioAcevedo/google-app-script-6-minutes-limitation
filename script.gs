function doGet() 
{
    var titleLabel, checkboxNameLabel, checkboxSurnamesLabel, checkboxEmailsLabel, checkboxPhonesLabel, checkboxAddressesLabel, checkboxGroupsLabel, checkboxNotesLabel, checkboxPositionLabel, checkboxCompaniesLabel, tableGroupsLabel;
    var originAccountLabel, processDescriptionLabel, processDescriptionStep1Label, processDescriptionStep2Label, processDescriptionStep3Label;
    var processNoCloseWindowWarningLabel, processReportButtonLabel, processGeneratingReportButtonLabel, noContactsProcessedLabel;
    
    try
    {
        switch(Session.getActiveUserLocale())
        {
            case "es":
                titleLabel = "Generador de informes de contactos basado en grupos";
                checkboxNameLabel = "Nombre";
                checkboxSurnamesLabel = "Apellidos";
                checkboxEmailsLabel = "Correos electrónicos";
                checkboxPhonesLabel = "Tel&eacute;fonos"
                checkboxAddressesLabel = "Direcciones"
                checkboxGroupsLabel = "Grupos (seleccionar solamente si son necesarios)"
                checkboxCompaniesLabel = "Compañías"
                checkboxPositionLabel = "Cargos"
                checkboxNotesLabel = "Notas"
                tableGroupsLabel = "Grupos"
                originAccountLabel = 'Cuenta origen: ' + Session.getActiveUser().getEmail()
                processDescriptionLabel = 'Este formulario genera un informe con los contactos pertenecientes al grupo que se selecciona incluyendo únicamente los campos marcados en el formulario.';
                processDescriptionStep1Label = "1. Seleccione los grupos que desea incluir en el informe"
                processDescriptionStep2Label = "2. Seleccione los campos que desea mostrar para cada contacto"
                processDescriptionStep3Label = "3. Pulse para generar el informe"
                processNoCloseWindowWarningLabel = "La operación puede tardar varios minutos. Por favor, no cierre esta ventana o el proceso de detendrá."
                processReportButtonLabel = "Generar Informe";
                processGeneratingReportButtonLabel = "Generando informe..."
                noContactsProcessedLabel = "Procesados 0 de 0 grupo(s) de contactos."
                break;
            default:
                titleLabel = "Report group-based generator";
                checkboxNameLabel = "Name";
                checkboxSurnamesLabel = "Surnames";
                checkboxEmailsLabel = "Emails";
                checkboxPhonesLabel = "Phones"
                checkboxAddressesLabel = "Addresses"
                checkboxGroupsLabel = "Groups (select only if needed)"
                checkboxCompaniesLabel = "Companies"
                checkboxPositionLabel = "Positions"
                checkboxNotesLabel = "Notes"
                tableGroupsLabel = "Groups"
                originAccountLabel = "Origin account: " + Session.getActiveUser().getEmail()
                processDescriptionLabel = "This form generates a report containing the contacts from selected group incluiding only selected fields.";
                processDescriptionStep1Label = "1. Select groups to be included."
                processDescriptionStep2Label = "2. Select fields to be showed for each contact,"
                processDescriptionStep3Label = "3. Press to generate the report"
                processNoCloseWindowWarningLabel = "This operation may take several minutes. Please, don't close this windows until the process finishes."
                processReportButtonLabel = "Create report";
                processGeneratingReportButtonLabel = "Creating report..."
                noContactsProcessedLabel = "0 of 0 contacts group(s) processed."
                break;
        }
        
        var app = UiApp.createApplication()
                       .setTitle(titleLabel)
     
        var form = app.createFormPanel().setId('panel');
        var flow = app.createFlowPanel().setId('flowPanel');
        
        var checkBoxName = app.createCheckBox()
        var checkBoxSurnames = app.createCheckBox()
        var checkBoxEmail = app.createCheckBox()
        var checkBoxPhone = app.createCheckBox()
        var checkBoxAddress = app.createCheckBox()
        var checkBoxGroups = app.createCheckBox()
        var checkBoxCompanies = app.createCheckBox()
        var checkBoxPositions = app.createCheckBox()
        var checkBoxNotes = app.createCheckBox()
      
        checkBoxName.setEnabled(true)
                    .setId("checkBoxName")
                    .setName("checkBoxName")
                    .setVisible(true)
                    .setValue(true)
                    .setHTML(checkboxNameLabel)
    
        checkBoxSurnames.setEnabled(true)
                        .setId("checkBoxSurnames")
                        .setName("checkBoxSurnames")
                        .setVisible(true)
                        .setValue(true)
                        .setHTML(checkboxSurnamesLabel)
    
        checkBoxEmail.setEnabled(true)
                     .setId("checkBoxEmail")
                     .setName("checkBoxEmail")
                     .setVisible(true)
                     .setValue(true)
                     .setHTML(checkboxEmailsLabel)
    
        checkBoxPhone.setEnabled(true)
                     .setId("checkBoxPhone")
                     .setName("checkBoxPhone")
                     .setTitle("Phone")
                     .setVisible(true)
                     .setValue(true)
                     .setHTML(checkboxPhonesLabel)
       
        checkBoxAddress.setEnabled(true)
                       .setId("checkBoxAddress")
                       .setName("checkBoxAddress")
                       .setVisible(true)
                       .setValue(true)
                       .setHTML(checkboxAddressesLabel)
        
        checkBoxGroups.setEnabled(true)
                       .setId("checkBoxGroups")
                       .setName("checkBoxGroups")
                       .setVisible(true)
                       .setValue(false)
                       .setHTML(checkboxGroupsLabel)
        
         checkBoxCompanies.setEnabled(true)
                         .setId("checkBoxCompanies")
                         .setName("checkBoxCompanies")
                         .setTitle("Companies")
                         .setVisible(true)
                         .setValue(false).setHTML(checkboxCompaniesLabel)
        
        checkBoxPositions.setEnabled(true)
                         .setId("checkBoxPositions")
                         .setName("checkBoxPositions")
                         .setTitle("Position")
                         .setVisible(true)
                         .setValue(false).setHTML(checkboxPositionLabel)
        
        checkBoxNotes.setEnabled(true)
                     .setId("checkBoxNotes")
                     .setName("checkBoxNotes")
                     .setTitle("Notes")
                     .setVisible(true)
                     .setValue(false).setHTML(checkboxNotesLabel)

        var contactGroups = ContactsApp.getContactGroups()
      
        var table = app.createFlexTable().setId("groupsTable");
    
        table.setStyleAttribute("margin", "10px")
        table.setStyleAttribute("border", "1px solid black")
        table.setCellPadding(3)
        table.setTitle(tableGroupsLabel)
      
        var NUMBER_OF_COLUMNS = 15;
            
        for(var i = 0 ; i < contactGroups.length/NUMBER_OF_COLUMNS ; i++)
        {
            for(var j = 0 ; j < NUMBER_OF_COLUMNS ; j++)
            {
              var contactIndex = (i*NUMBER_OF_COLUMNS)+j;
              
              var group = contactGroups[contactIndex];
              
              if(group && !group.isSystemGroup())
              {
                  var checkBox = app.createCheckBox(group.getName())
                                      .setId("group" + contactIndex)
                                      .setName("group" + contactIndex)
              
                  var checkBoxId = app.createTextBox()
                                        .setValue(group.getId())
                                        .setId("group_"+ contactIndex +"_id")
                                        .setName("group_"+ contactIndex +"_id")
                                        .setVisible(false)
    
                  table.setWidget((j+1), (i*NUMBER_OF_COLUMNS)+1, checkBox); 
                  flow.add(checkBoxId)
              }
    
            }
        }
        
        var numberOfGroupsTextBox = app.createTextBox()
                                       .setId("numberOfGroupsTextBox")
                                       .setName("numberOfGroupsTextBox")
                                       .setValue(contactGroups.length)
                                       .setVisible(false)
        
    
        flow.setTitle(titleLabel);
        
        flow.add(app.createLabel(processDescriptionLabel)
                 .setStyleAttributes({fontSize : "20px", fontWeight : "bold", margin: "10px"}));
        
        flow.add(app.createLabel(originAccountLabel).setStyleAttributes({fontWeight : "bold", margin: "10px"}));
    
        flow.add(app.createLabel(processDescriptionStep1Label).setStyleAttributes({fontWeight : "bold", margin: "10px"}))
        
        flow.add(table)
    
        flow.add(app.createLabel(processDescriptionStep2Label).setStyleAttributes({fontWeight : "bold", margin: "10px"}))
    
        var fieldsTable = app.createFlexTable().setId("fieldsTable");
    
        fieldsTable.setStyleAttribute("margin", "10px")
        fieldsTable.setStyleAttribute("border", "1px solid black")
        fieldsTable.setCellPadding(3)
        fieldsTable.setTitle("Groups")
        
        fieldsTable.setWidget(0, 0, checkBoxName)
        fieldsTable.setWidget(0, 1,checkBoxSurnames)
        fieldsTable.setWidget(0, 2, checkBoxAddress)
        fieldsTable.setWidget(1, 0, checkBoxEmail)
        fieldsTable.setWidget(1, 1, checkBoxPhone)
        fieldsTable.setWidget(1, 2, checkBoxGroups)
        fieldsTable.setWidget(2, 0, checkBoxCompanies)
        fieldsTable.setWidget(2, 1, checkBoxPositions)
        fieldsTable.setWidget(2, 2, checkBoxNotes)
    
        var spreadSheetFileIdTextBox = app.createTextBox()
                                       .setId("spreadSheetFileIdTextBox")
                                       .setName("spreadSheetFileIdTextBox")
                                       .setVisible(false)
                                       
        flow.add(spreadSheetFileIdTextBox)
        
        flow.add(fieldsTable);
      
        flow.add(numberOfGroupsTextBox);
            
        flow.add(app.createLabel(processDescriptionStep3Label).setStyleAttributes({fontWeight : "bold", margin: "10px"}))
    
        var infoLabel = app.createLabel(processNoCloseWindowWarningLabel).setVisible(false).setStyleAttributes({fontWeight : "bold", margin: "10px"})
        
        var submitLabelChangeHandler = app.createClientHandler()
                                          .forEventSource().setHTML(processGeneratingReportButtonLabel)
                                          .forTargets(infoLabel).setVisible(true)
        
        var handlerSubmit = app.createServerHandler('prepareData').addCallbackElement(flow);
        flow.add(app.createButton(processReportButtonLabel, handlerSubmit).setId("submitButton").setStyleAttributes({fontWeight : "bold", margin: "10px"}).addClickHandler(submitLabelChangeHandler));
    
        flow.add(infoLabel)
        flow.add(app.createLabel(noContactsProcessedLabel).setId("remainingGroups").setVisible(false).setStyleAttributes({fontWeight : "bold", margin: "10px"}))
  
        form.add(flow);
    
        app.add(form);
    }
    catch(e)
    {
        GmailApp.sendEmail( "carlos.delgado@proyecti.es,antonio.acevedo@proyecti.es", 
                            "Error - Informe sobre contactos", 
                            "Error generando informe de contactos:" + JSON.stringify(e));

        return showReportGenerationResult(app, false)
    }
    
    return app;  
}

function prepareData(eventInfo)
{
    var app = UiApp.getActiveApplication()
    
    app.getElementById("submitButton").setEnabled(false)

    //I. Buscamos los grupos que ha solicitado el usuario y los almacenamos en un array
    var numberOfGroups = eventInfo.parameter.numberOfGroupsTextBox;

    var requiredGroupsIds = []
    for(var i = 0 ; i < numberOfGroups ; i++)
    {
        var selected = eventInfo.parameter["group"+i]
        if(selected == "true")
        {
            requiredGroupsIds.push(eventInfo.parameter["group_"+i+"_id"])
        }
    }
    
    var remainingGroupsLabel;
    var spreadSheetFile = createTemporalSpreadSheet();
    app.getElementById("spreadSheetFileIdTextBox").setValue(spreadSheetFile.getId())

    switch(Session.getActiveUserLocale())
    {
            case "es":
                remainingGroupsLabel = "Procesados 0 de " + requiredGroupsIds.length + " grupo(s) de contactos."
                break;
            default:
                remainingGroupsLabel = "0 of " + requiredGroupsIds.length + " contacts group(s) processed."
                break;
    }
   
    //I. Meter un checkbox que tenga un handler que llame a la funcion de procesar un grupo
    var flowPanel = app.getElementById("flowPanel")
    var processGroupHandler = app.createServerHandler('processGroup').addCallbackElement(flowPanel);
    var loopCheckbox = app.createCheckBox('loopCheckbox').addValueChangeHandler(processGroupHandler).setId('loopCheckbox').setVisible(false)

    //II. Crear un textbox que vaya almacenando el array de grupos que quedan por ser procesados
    var numberOfSelectedGroups = app.createTextBox().setId("numberOfSelectedGroups").setVisible(false).setName('numberOfSelectedGroups').setValue(requiredGroupsIds.length)
    var requiredGroupsIdsTextBox = app.createTextBox().setId("requiredGroupsIdsTextBox").setVisible(false).setName('requiredGroupsIdsTextBox').setValue(JSON.stringify(requiredGroupsIds))
        
    flowPanel.add(numberOfSelectedGroups)
    flowPanel.add(loopCheckbox)
    flowPanel.add(requiredGroupsIdsTextBox)

    app.getElementById("remainingGroups").setVisible(true).setText(remainingGroupsLabel);

    //II. Lanzamos el bucle
    
    loopCheckbox.setValue(true,true)
    
    return app;
}

//http://stackoverflow.com/questions/16886550/what-happens-when-i-sleep-in-gas-execution-time-limit-workaround
function processGroup(eventInfo)
{
    try
    {
        var app = UiApp.getActiveApplication();
      
        var needName = eventInfo.parameter.checkBoxName == "true";
        var needSurnames = eventInfo.parameter.checkBoxSurnames == "true";
        var needEmails = eventInfo.parameter.checkBoxEmail == "true";
        var needAddresses = eventInfo.parameter.checkBoxAddress == "true";
        var needPhones = eventInfo.parameter.checkBoxPhone == "true";
        var needGroups = eventInfo.parameter.checkBoxGroups == "true";
        var needCompanies = eventInfo.parameter.checkBoxCompanies == "true";
        var needPositions = eventInfo.parameter.checkBoxPositions == "true";
        var needNotes = eventInfo.parameter.checkBoxNotes == "true";
        var spreadSheetFileId = eventInfo.parameter.spreadSheetFileIdTextBox;
        var requiredGroupsIds = JSON.parse(eventInfo.parameter.requiredGroupsIdsTextBox)
        var numberOfSelectedGroups = eventInfo.parameter.numberOfSelectedGroups

        var nameHeader, surnamesHeader, email1Header, email2Header, phone1Header, phone2Header, phone3Header, phone4Header, address1Header, address2Header, groupsHeader, company1Header, company2Header, position1Header, position2Header, notesHeader;
        var numberOfProcessedGroupsLabel, newFileNamePrefix, subjectEmail, bodyEmailPrefix
        
        switch(Session.getActiveUserLocale())
        {
            case "es":
                nameHeader = "Nombre"
                surnamesHeader = "Apellidos"
                email1Header = "Correo electrónico 1"
                email2Header = "Correo electrónico 2"
                phone1Header = "Teléfono 1"
                phone2Header = "Teléfono 2"
                phone3Header = "Teléfono 3"
                phone4Header = "Teléfono 4"
                address1Header = "Dirección 1"
                address2Header = "Dirección 2"
                groupsHeader = "Grupos"
                company1Header = "Compañía 1"
                company2Header = "Compañía 2"
                position1Header = "Position 1"
                position2Header = "Position 2"
                notesHeader = "Notas"
                numberOfProcessedGroupsLabel = "Procesados " + (numberOfSelectedGroups-requiredGroupsIds.length) + " de " + numberOfSelectedGroups + " grupo(s) de contactos."
                newFileNamePrefix = "Informe de contactos"
                subjectEmail = "Informe generado - contactos por grupo";
                bodyEmailPrefix = "Informe de contactos generado satisfactoriamente. Para acceder al mismo pinche sobre el siguiente link: ";
                break;
                
            default:
                nameHeader = "Name"
                surnamesHeader = "Surnames"
                email1Header = "Email 1"
                email2Header = "Email 2"
                phone1Header = "Phone 1"
                phone2Header = "Phone 2"
                phone3Header = "Phone 3"
                phone4Header = "Phone 4"
                address1Header = "Address 1"
                address2Header = "Address 2"
                company1Header = "Company 1"
                company2Header = "Company 2"
                position1Header = "Position 1"
                position2Header = "Position 2"
                notesHeader = "Notes"
                groupsHeader = "Group(s)"
                numberOfProcessedGroupsLabel = (numberOfSelectedGroups-requiredGroupsIds.length) + " of " + numberOfSelectedGroups + " contacts group(s) processed."
                newFileNamePrefix = "Contacts report"
                subjectEmail = "Report created - Contacts per group";
                bodyEmailPrefix = "Contacts per group report created successfully. It can be accessed through the following link: ";
                break;
        }
        
        //Save the data into the spreadsheet
        var spreadSheet =  SpreadsheetApp.openById(spreadSheetFileId);
    
        //Si no quedan elementos en el array, avisamos al usuario y cortamos
        if(requiredGroupsIds.length == 0)
        {
            return prepareFinalSpreadSheet(spreadSheet)
        }
    
        var sheet = spreadSheet.getActiveSheet()
            
        var currentLastRow = sheet.getLastRow()+1
    
        var contacts;
        
        //I. Create the header based on user selection
        var headers = []
        var contactInfoTemplate = []
              
        if(needName)
        {
            contactInfoTemplate.push('')
            headers.push(nameHeader)
        }
          
        if(needSurnames)
        {
            contactInfoTemplate.push('')
            headers.push(surnamesHeader)
        }
        
         if(needEmails)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push(email1Header)
            headers.push(email2Header)
        }
          
        if(needPhones)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push(phone1Header)
            headers.push(phone2Header)
            headers.push(phone3Header)
            headers.push(phone4Header)
        }
          
        if(needAddresses)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push(address1Header)
            headers.push(address2Header)
        }
        
        if(needCompanies && !needPositions)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push(company1Header)
            headers.push(company2Header)
        }
        else if(needCompanies && needPositions)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push(company1Header)
            headers.push(position1Header)
            headers.push(company2Header)
            headers.push(position2Header)
        }
        else if(needPositions)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push(position1Header)
            headers.push(position2Header)
        }
        
        if(needGroups)
        {
            contactInfoTemplate.push('')
            headers.push(groupsHeader)
        }
        
        if(needNotes)
        {
            contactInfoTemplate.push('')
            headers.push(notesHeader)
        }
        
        var allContactsInfo = []
        var headerRowsFileNumber = []
        var groupTitleRowsFileNumber = []
            
        //II. Get first group in the array
            
        var contactGroup = ContactsApp.getContactGroupById(requiredGroupsIds.shift())
        var contacts = contactGroup.getContacts()
        var result;
    
        var groupTitleRow = contactInfoTemplate.slice(0);
        groupTitleRow[0] = contactGroup.getName()
        allContactsInfo.push(contactInfoTemplate.slice(0))
        allContactsInfo.push(groupTitleRow)
        groupTitleRowsFileNumber.push(sheet.getLastRow() + allContactsInfo.length)
                
        allContactsInfo.push(headers)
        headerRowsFileNumber.push(sheet.getLastRow() + allContactsInfo.length)
                
        allContactsInfo.push(contactInfoTemplate.slice(0))
                
        var numberOfContactsInGroup = contacts.length;
        
        //IV. Iterate over each group contact
        //IV. Iterate over each group contact
        for (var j=0; j < contacts.length; j++) 
        {
            Logger.log("== Añadiendo contacto == ")
            var contact = contacts[j];
     
            var contactInfo = []
                                    
            //V. Adding contact name
            if(needName)
            {
                Logger.log("Name: " + contact.getGivenName())
                contactInfo.push(contact.getGivenName())
            }
              
            //VI. Adding contact surnames
            if(needSurnames)
            {
                Logger.log("Surname: " + contact.getFamilyName())
                contactInfo.push(contact.getFamilyName())
            }
            
            //VII. Adding contact emails
            if(needEmails)
            {
                var email1, email2;
                    
                switch(contact.getEmails().length)
                {
                    case 0: 
                        email1 = "";
                        email2 = "";
                        break;
                    case 1:
                        email1 = contact.getEmails()[0].getAddress();
                        email2 = "";
                        break;
                    default:                    
                        var primaryEmail = contact.getEmails()[0].isPrimary()
                        email1 = primaryEmail ? contact.getEmails()[0].getAddress() : contact.getEmails()[1].getAddress();
                        email2 = primaryEmail ? contact.getEmails()[1].getAddress() : contact.getEmails()[0].getAddress();
                        break;
                }      
                    
                contactInfo.push(email1)
                contactInfo.push(email2)
            }
              
            //VIII. Adding contact phones
            if(needPhones)
            {
                var phone1, phone2, phone3, phone4;
                  
                switch(contact.getPhones().length)
                {
                    case 0: 
                        phone1 = "";
                        phone2 = "";
                        phone3 = "";
                        phone4 = "";
                        break;
                        
                    case 1:
                        phone1 = contact.getPhones()[0].getPhoneNumber();
                        phone2 = "";
                        phone3 = "";
                        phone4 = "";
                        break;
                        
                    case 2:
                        var primaryPhone = contact.getPhones()[0].isPrimary()
                        phone1 = primaryPhone ? contact.getPhones()[0].getPhoneNumber() : contact.getPhones()[1].getPhoneNumber();
                        phone2 = primaryPhone ? contact.getPhones()[1].getPhoneNumber() : contact.getPhones()[0].getPhoneNumber();
                        phone3 = "";
                        phone4 = "";
                        break;
                    
                    case 3:
                        phone1 = contact.getPhones()[0].getPhoneNumber()
                        phone2 = contact.getPhones()[1].getPhoneNumber()
                        phone3 = contact.getPhones()[2].getPhoneNumber()
                        phone4 = "";
                        break;
                    
                    default:
                        phone1 = contact.getPhones()[0].getPhoneNumber()
                        phone2 = contact.getPhones()[1].getPhoneNumber()
                        phone3 = contact.getPhones()[2].getPhoneNumber()
                        phone4 = contact.getPhones()[3].getPhoneNumber()
                        break;
                }      
                  
                contactInfo.push(phone1)
                contactInfo.push(phone2)
                contactInfo.push(phone3)
                contactInfo.push(phone4)
            }
          
            //IX. Adding contact addresses
            if(needAddresses)
            {
                var address1, address2;
              
                switch(contact.getAddresses().length)
                {
                    case 0: 
                        address1 = "";
                        address2 = "";
                        break;
                        
                    case 1:
                        address1 = contact.getAddresses()[0].getAddress();
                        address2 = "";
                        break;
                        
                    default:
                        var primaryAddress = contact.getAddresses()[0].isPrimary()
                        address1 = primaryAddress ? contact.getAddresses()[0].getAddress() : contact.getAddresses()[1].getAddress();
                        address2 = primaryAddress ? contact.getAddresses()[1].getAddress() : contact.getAddresses()[0].getAddress();
                        break;
                }
              
                contactInfo.push(address1)
                contactInfo.push(address2)
            }
            
            //X. Adding contact companies
            if(needCompanies || needPositions)
            {
                var company1, company2, position1, position2;
                    
                switch(contact.getCompanies().length)
                {
                    case 0: 
                        company1 = "";
                        position1 = "";
                        company2 = "";
                        position2 = "";
                        break;
                    case 1:
                        company1 = (needCompanies) ? contact.getCompanies()[0].getCompanyName() : "";
                        position1 = (needPositions) ? contact.getCompanies()[0].getJobTitle() : "";
                        company2 = "";
                        position2 = "";
                        break;
                    default:                    
                        var primaryCompany = contact.getCompanies()[0].isPrimary()
                        company1 = (needCompanies) ? (primaryEmail ? contact.getCompanies()[0].getCompanyName() : contact.getCompanies()[1].getCompanyName()) : "";
                        position1 = (needPositions) ? (primaryEmail ? ((needPositions) ? contact.getCompanies()[0].getJobTitle() : "") : ((needPositions) ? contact.getCompanies()[1].getJobTitle() : "")) : "";
                        company2 = (needCompanies) ? (primaryEmail ? contact.getCompanies()[1].getCompanyName() : contact.getCompanies()[0].getCompanyName()) : "";
                        position2 = (needPositions) ? (primaryEmail ? ((needPositions) ? contact.getCompanies()[1].getJobTitle() : "") : ((needPositions) ? contact.getCompanies()[0].getJobTitle() : "")) : "";
                        break;
                }      
                
                if(needCompanies && needPositions)
                {
                    contactInfo.push(company1)
                    contactInfo.push(position1)
                    contactInfo.push(company2)
                    contactInfo.push(position2)
                }
                else if(!needPositions)
                {
                    contactInfo.push(company1)
                    contactInfo.push(company2)
                }
                else
                {
                    contactInfo.push(position1)
                    contactInfo.push(position2)
                }
            }
              
              
            //XI. Adding contact groups
            if (needGroups)
            {
                var groups = contact.getContactGroups()
                var groupsAsArray = []
              
                for (var k = 0 ; k < groups.length; k++) 
                {
                    groupsAsArray.push(groups[k].getName())
                }
              
                contactInfo.push(groupsAsArray.sort().join())
            }
          
            if(needNotes)
            {
                contactInfo.push(contact.getNotes())            
            }
            
            allContactsInfo.push(contactInfo)
        }
    
        allContactsInfo.push(contactInfoTemplate.slice(0))
    
        sheet.insertRowsAfter(currentLastRow, allContactsInfo.length);
        
        sheet.getRange(currentLastRow, 1, allContactsInfo.length, headers.length).setValues(allContactsInfo);
        
        for(var i = 0 ; i < headerRowsFileNumber.length ; i++)
        {
            sheet.getRange(headerRowsFileNumber[i], 1, 1, headers.length).setFontWeight("bold").setBackground("blue").setFontColor("white").setHorizontalAlignment("center")
        }
         
        for(var i = 0 ; i < groupTitleRowsFileNumber.length ; i++)
        {
            sheet.getRange(groupTitleRowsFileNumber[i], 1, 1, 1).setFontWeight("bold").setHorizontalAlignment("center").setFontLine("underline")      
        }
        
        app.getElementById("requiredGroupsIdsTextBox").setValue(JSON.stringify(requiredGroupsIds))
    
        var numberOfGroups = eventInfo.parameter.numberOfGroupsTextBox
    
        app.getElementById("remainingGroups").setVisible(true).setText(numberOfProcessedGroupsLabel);
        app.getElementById('loopCheckbox').setValue(false,false)
        app.getElementById('loopCheckbox').setValue(true,true)
        
        return app;
    }
    catch(e)
    {
        GmailApp.sendEmail( "carlos.delgado@proyecti.es,antonio.acevedo@proyecti.es", 
                            "Error - Contacts report", 
                           "Error generating contacts report:" + JSON.stringify(e));
        
        return showReportGenerationResult(app, false)
    }
}


function showReportGenerationResult(app, success, spreadSheetUrl)
{
    var form = app.createFormPanel().setId('panel');
    var flow = app.createFlowPanel();
  
    var flowTitle, correctReportLabel, correctReportLinkLabel, correctReportLink, correctReportEmailSentLabel, errorReportLabel, errorReportSupportLabel
    
    switch(Session.getActiveUserLocale())
    {
        case "es":
            flowTitle = 'Generador de informes de contactos basado en grupos.'
            correctReportLabel = 'Informe generado correctamente.'
            correctReportLinkLabel = 'Puede acceder al informe desde el siguiente enlace:'
            correctReportLink = "Acceder al informe"
            correctReportEmailSentLabel = 'Adicionalmente se ha enviado a su bandeja de entrada un correo para consultas posteriores'
            errorReportLabel = 'El informe no ha podido ser generado.'
            errorReportSupportLabel = 'Se ha enviado un correo informando de la incidencia al responsable técnico para solucionarlo a la mayor brevedad posible. Disculpen las molestias'
            break;
        
        default:
            flowTitle = 'Generador de informes de contactos basado en grupos.'
            correctReportLabel = 'Report created successfully.'
            correctReportLinkLabel = 'It can be accessed throught the following link:'
            correctReportLink = "Access to the report"
            correctReportEmailSentLabel = 'In addition, an email has been sent to your inbox containing this link'
            errorReportLabel = 'The report cannot be created'
            errorReportSupportLabel = 'An email has been sent to support team with the incidence. It will be fixed as soon as possible. Sorry for the inconvenience'
            break;
    }

    flow.setTitle(flowTitle);
    
    if(success)
    {
        flow.add(app.createLabel(correctReportLabel).setStyleAttributes({fontSize : "20px", fontWeight : "bold", margin: "10px"}));
    
        flow.add(app.createInlineLabel(correctReportLinkLabel).setStyleAttributes({margin: "10px"}));
        flow.add(app.createAnchor(correctReportLink, spreadSheetUrl).setStyleAttributes({margin: "10px"}));

        flow.add(app.createLabel(correctReportEmailSentLabel).setStyleAttributes({margin: "10px"}));
    }
    else
    {
        flow.add(app.createLabel(errorReportLabel).setStyleAttributes({fontSize : "20px", fontWeight : "bold", margin: "10px"}));
    
        flow.add(app.createInlineLabel(errorReportSupportLabel).setStyleAttributes({margin: "10px"}));  
    }
    
    form.add(flow);
  
    app.add(form);
    
    return app;
}


function createTemporalSpreadSheet()
{
    var folder = getReportsTemporalFolder();
    var temporalName;
  
    switch(Session.getActiveUserLocale())
    {
        case "es":
            temporalname = "Informe de contactos (" + new Date() + ")";
            break;
        default:
            temporalname = "Contacts report (" + new Date() + ")";
            break;
    }
    
    var temporalSpreadSheet
    
    try
    {
        temporalSpreadSheet = SpreadsheetApp.create(temporalName);
        var temporalFile = DriveApp.getFileById(temporalSpreadSheet.getId());

        var file = temporalFile.makeCopy(temporalName,folder);
        DriveApp.getRootFolder().removeFile( temporalFile );
    }
    catch(e)
    {
        temporalSpreadSheet = getSpreadSheetTemplate(temporalName)
    }
    return SpreadsheetApp.openById(temporalSpreadSheet.getId());
}


function getReportsFolder()
{
    var folderName;
    var tmpFolderName = "tmp";
    
    switch(Session.getActiveUserLocale())
    {
        case "es":
            folderName = "informes-contactos";
            break;
        default:
            folderName = "contacts-report";
            break;
    }
    
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;

    while (folders.hasNext()) 
    {
        folder = folders.next();
    }

    if(folder == null)
    {
        folder = DriveApp.createFolder(folderName);
        folder.createFolder(tmpFolderName)
    }
    return folder;
}

function getReportsTemporalFolder()
{
    var tmpFolderName = "tmp";

    var reportFolder = getReportsFolder()
    
    var folders = reportFolder.getFoldersByName(tmpFolderName);
    var temporalFolder;
    
    while(folders.hasNext())
    {
        temporalFolder = folders.next();
    }
    
    return temporalFolder;
}

function prepareFinalSpreadSheet(spreadSheet)
{
    var reportName, correctReportSubject, correctReportBody, correctReportLink
    switch(Session.getActiveUserLocale())
    {
        case "es":
            reportName = "Informe de contactos (" + new Date() + ")"
            correctReportSubject = "Informe de contactos por grupo generado"
            correctReportBody = 'Informe de contactos por grupo generado satisfactoriamente.'
            correctReportLink = "Acceder al informe"
            break;
        
        default:
            reportName = "Informe de contactos (" + new Date() + ")"
            correctReportSubject = "Contacts report generated"
            correctReportBody = 'Report created successfully.'
            correctReportLink = "Access to the report"
            break;
    }


    var file = DriveApp.getFileById(spreadSheet.getId())
    
    var newFile = file.makeCopy(reportName, getReportsFolder())
    
    file.setTrashed(true);
    
    var body = correctReportBody + "<br /> <a href='" + newFile.getUrl() + "'>" + correctReportLink + "</a>";
     
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), correctReportSubject, "", { htmlBody : body} )
    
    return showReportGenerationResult(UiApp.getActiveApplication(), true, newFile.getUrl())
}

function getSpreadSheetTemplate(fileName)
{
    var automationsFolderName, contactsScriptFolder, contactsSpreadSheetTemplateName = "template";
    
    switch(Session.getActiveUserLocale())
    {
        case "es":
            automationsFolderName = "automatizaciones-NO TOCAR";
            contactsScriptFolder = "Informe de contactos por grupo"
            break;
        default:
            automationsFolderName = "automations-NOT TOUCH";
            contactsScriptFolder = "Groups report"
            break;
    }
    
    var folders = DriveApp.getRootFolder().getFoldersByName(automationsFolderName)
    
    var automationsFolder
    
    while(folders.hasNext())
    {
        automationsFolder = folders.next()
    }
    
    if(automationsFolder == null)
    {
        automationsFolder = DriveApp.getRootFolder().createFolder(automationsFolderName)
    }
    
    var infolders = automationsFolder.getFoldersByName(contactsScriptFolder);

    var templatesFolder;

    while(infolders.hasNext())
    {
        templatesFolder = infolders.next()
    }

    var spreadSheetTemplate;
    
    if(templatesFolder == null)
    {
        templatesFolder = automationsFolder.createFolder(contactsScriptFolder)
        spreadSheetTemplate = SpreadsheetApp.create(contactsSpreadSheetTemplateName);
        var temporalFile = DriveApp.getFileById(spreadSheetTemplate.getId());

        var file = temporalFile.makeCopy("template",templatesFolder);
        DriveApp.getRootFolder().removeFile(temporalFile);
        return spreadSheetTemplate; 
    }
    else
    {
        spreadSheetTemplate = templatesFolder.getFilesByName(contactsSpreadSheetTemplateName).next()
    }
    
    return temporalSpreadSheet = spreadSheetTemplate.makeCopy(fileName, getReportsTemporalFolder())
}
//http://stackoverflow.com/questions/19408464/apps-script-oauth2-youtube-api-v3
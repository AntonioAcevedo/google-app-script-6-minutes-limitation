function doGet() 
{
    try
    {
        var app = UiApp.createApplication()
                     .setTitle("Generador de informes de contactos basado en grupos")
     
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
                    .setTitle("Name")
                    .setVisible(true)
                    .setValue(true).setHTML("Nombre")
    
        checkBoxSurnames.setEnabled(true)
                        .setId("checkBoxSurnames")
                        .setName("checkBoxSurnames")
                        .setTitle("Surnames")
                        .setVisible(true)
                        .setValue(true).setHTML("Apellido")
    
        checkBoxEmail.setEnabled(true)
                     .setId("checkBoxEmail")
                     .setName("checkBoxEmail")
                     .setTitle("Email")
                     .setVisible(true)
                     .setValue(true).setHTML("Email")
    
        checkBoxPhone.setEnabled(true)
                     .setId("checkBoxPhone")
                     .setName("checkBoxPhone")
                     .setTitle("Phone")
                     .setVisible(true)
                     .setValue(true).setHTML("Tel&eacute;fono")
       
        checkBoxAddress.setEnabled(true)
                       .setId("checkBoxAddress")
                       .setName("checkBoxAddress")
                       .setTitle("Address")
                       .setVisible(true)
                       .setValue(true).setHTML("Direcci&oacute;n")
        
        checkBoxGroups.setEnabled(true)
                       .setId("checkBoxGroups")
                       .setName("checkBoxGroups")
                       .setTitle("Groups")
                       .setVisible(true)
                       .setValue(false).setHTML("Grupos (No marcar si no es completamente necesario)")
        
        checkBoxCompanies.setEnabled(true)
                         .setId("checkBoxCompanies")
                         .setName("checkBoxCompanies")
                         .setTitle("Companies")
                         .setVisible(true)
                         .setValue(false).setHTML("Empresa")
        
        checkBoxPositions.setEnabled(true)
                         .setId("checkBoxPositions")
                         .setName("checkBoxPositions")
                         .setTitle("Position")
                         .setVisible(true)
                         .setValue(false).setHTML("Cargo")
        
        checkBoxNotes.setEnabled(true)
                     .setId("checkBoxNotes")
                     .setName("checkBoxNotes")
                     .setTitle("Notes")
                     .setVisible(true)
                     .setValue(false).setHTML("Notas")
        
        var contactGroups = ContactsApp.getContactGroups()
      
        var table = app.createFlexTable().setId("groupsTable");
    
        table.setStyleAttribute("margin", "10px")
        table.setStyleAttribute("border", "1px solid black")
        table.setCellPadding(3)
        table.setTitle("Groups")
      
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
        
    
        flow.setTitle('Generador de informes de contactos basado en grupos.');
        
        flow.add(app.createLabel('Este formulario genera un informe con los contactos pertenecientes al grupo que se selecciona incluyendo únicamente los campos marcados en el formulario.')
                 .setStyleAttributes({fontSize : "20px", fontWeight : "bold", margin: "10px"}));
        
        flow.add(app.createLabel('Cuenta origen: ' + Session.getActiveUser().getEmail()).setStyleAttributes({fontWeight : "bold", margin: "10px"}));
    
        flow.add(app.createLabel("1. Seleccione los grupos que desea incluir en el informe").setStyleAttributes({fontWeight : "bold", margin: "10px"}))
        
        flow.add(table)
    
        flow.add(app.createLabel("2. Seleccione los campos que desea mostrar para cada contacto").setStyleAttributes({fontWeight : "bold", margin: "10px"}))
    
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
            
        flow.add(app.createLabel("3. Pulse para generar el informe").setStyleAttributes({fontWeight : "bold", margin: "10px"}))
    
        var infoLabel = app.createLabel("La operación puede tardar varios minutos. Por favor, no cierre esta ventana o el proceso de detendrá.").setVisible(false).setStyleAttributes({fontWeight : "bold", margin: "10px"})
        
        var submitLabelChangeHandler = app.createClientHandler()
                                          .forEventSource().setHTML("Generando informe...")
                                          .forTargets(infoLabel).setVisible(true)
        
        var handlerSubmit = app.createServerHandler('prepareData').addCallbackElement(flow);
        flow.add(app.createButton("Generar Informe",handlerSubmit).setId("submitButton").setStyleAttributes({fontWeight : "bold", margin: "10px"}).addClickHandler(submitLabelChangeHandler));
    
        flow.add(infoLabel)
        flow.add(app.createLabel("Procesados 0 de 0 grupo(s) de contactos.").setId("remainingGroups").setVisible(false).setStyleAttributes({fontWeight : "bold", margin: "10px"}))
    
        form.add(flow);
    
        app.add(form);
    
        Logger.log("Creado spreadsheet dentro de la carpeta")

    }
    catch(e)
    {
        GmailApp.sendEmail( "carlos.delgado@proyecti.es,antonio.acevedo@proyecti.es", 
                            "Error - Informe sobre contactos", 
                            "Error generando informe de contactos:" + JSON.stringify(e));
        
        Logger.log("Error encontrado: ", JSON.stringify(e));
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

    var spreadSheetFile = createTemporalSpreadSheet();
    app.getElementById("spreadSheetFileIdTextBox").setValue(spreadSheetFile.getId())

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

    app.getElementById("remainingGroups").setVisible(true).setText("Procesados 0 de " + requiredGroupsIds.length + " grupo(s) de contactos.");

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
            headers.push("Nombre")
        }
          
        if(needSurnames)
        {
            contactInfoTemplate.push('')
            headers.push("Apellidos")
        }
        
         if(needEmails)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push("Email 1")
            headers.push("Email 2")
        }
          
        if(needPhones)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push("Teléfono 1")
            headers.push("Teléfono 2")
            headers.push("Teléfono 3")
            headers.push("Teléfono 4")
        }
          
        if(needAddresses)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push("Dirección 1")
            headers.push("Dirección 2")
        }
        
        if(needCompanies && !needPositions)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push("Empresa 1")
            headers.push("Empresa 2")
        }
        else if(needCompanies && needPositions)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push("Empresa 1")
            headers.push("Cargo 1")
            headers.push("Empresa 1")
            headers.push("Cargo 2")
        }
        else if(needPositions)
        {
            contactInfoTemplate.push('')
            contactInfoTemplate.push('')
            headers.push("Cargo 1")
            headers.push("Cargo 2")
        }
        
        if(needGroups)
        {
            contactInfoTemplate.push('')
            headers.push("Grupo/s")
        }
      
        if(needNotes)
        {
            contactInfoTemplate.push('')
            headers.push("Notas")
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
    
        app.getElementById("remainingGroups").setVisible(true).setText("Procesados " + (numberOfSelectedGroups-requiredGroupsIds.length) + " de " + numberOfSelectedGroups + " grupo(s) de contactos.");
        app.getElementById('loopCheckbox').setValue(false,false)
        app.getElementById('loopCheckbox').setValue(true,true)
        
        return app;
    }
    catch(e)
    {
        GmailApp.sendEmail( "carlos.delgado@proyecti.es,antonio.acevedo@proyecti.es", 
                            "Error - Informe sobre contactos", 
                           "Error generando informe de contactos:" + JSON.stringify(e));
        
        Logger.log("Error encontrado: ", JSON.stringify(e));
        return showReportGenerationResult(app, false)
    }
}

function showReportGenerationResult(app, success, spreadSheetUrl)
{
    var form = app.createFormPanel().setId('panel');
    var flow = app.createFlowPanel();
  
    flow.setTitle('Generador de informes de contactos basado en grupos.');
    
    if(success)
    {
        flow.add(app.createLabel('Informe generado correctamente.').setStyleAttributes({fontSize : "20px", fontWeight : "bold", margin: "10px"}));
    
        flow.add(app.createInlineLabel('Puede acceder al informe desde el siguiente enlace:').setStyleAttributes({margin: "10px"}));
        flow.add(app.createAnchor("Acceder al informe", spreadSheetUrl).setStyleAttributes({margin: "10px"}));

        flow.add(app.createLabel('Adicionalmente se ha enviado a su bandeja de entrada un correo para consultas posteriores').setStyleAttributes({margin: "10px"}));
    }
    else
    {
        flow.add(app.createLabel('El informe no ha podido ser generado.').setStyleAttributes({fontSize : "20px", fontWeight : "bold", margin: "10px"}));
    
        flow.add(app.createInlineLabel('Se ha enviado un correo informando de la incidencia al responsable técnico para solucionarlo a la mayor brevedad posible. Disculpen las molestias').setStyleAttributes({margin: "10px"}));  
    }
    
    form.add(flow);
  
    app.add(form);
    
    return app;
}

function createTemporalSpreadSheet()
{
    var folder = getReportsTemporalFolder();
  
    var temporalName = "Informe de contactos (" + new Date() + ")";
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
    var folders = DriveApp.getFoldersByName('informes-contactos');
    var folder;

    while (folders.hasNext()) 
    {
        folder = folders.next();
    }

    if(folder == null)
    {
        folder = DriveApp.createFolder("informes-contactos");
        folder.createFolder("tmp")
    }
    return folder;
}

function getReportsTemporalFolder()
{
    var reportFolder = getReportsFolder()
    
    var folders = reportFolder.getFoldersByName("tmp");
    var temporalFolder;
    
    while(folders.hasNext())
    {
        temporalFolder = folders.next();
    }
    
    return temporalFolder;
}

function prepareFinalSpreadSheet(spreadSheet)
{
    var app = UiApp.getActiveApplication()
    var file = DriveApp.getFileById(spreadSheet.getId())
    
    var newFile = file.makeCopy(file.getName(), getReportsFolder())
    
    file.setTrashed(true);
 
    var body = "Informe de contactos por grupo generado satisfactoriamente. <br /> <a href='" + newFile.getUrl() + "'>Acceder al informe</a>";
     
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), "Informe de contactos por grupo generado", "", { htmlBody : body} )
    
    return showReportGenerationResult(app, true, newFile.getUrl())
}

function getSpreadSheetTemplate(fileName)
{
    var folders = DriveApp.getRootFolder().getFoldersByName("automatizaciones-NO TOCAR")
    
    var automationsFolder
    
    while(folders.hasNext())
    {
        automationsFolder = folders.next()
    }
    
    var infolders = automationsFolder.getFoldersByName("Informe de contactos por grupo");

    var templatesFolder;

    while(infolders.hasNext())
    {
        templatesFolder = infolders.next()
    }

    var contactsPerGroupReportSpreadsheetTemplate;

    return templatesFolder.getFilesByName("template").next().makeCopy(fileName, getReportsTemporalFolder())
}
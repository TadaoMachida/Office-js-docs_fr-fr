
# Manifestes des compléments Outlook

Un complément Outlook se compose de deux éléments : le manifeste de complément XML et une page web, pris en charge par la bibliothèque JavaScript pour les Compléments Office (office.js). Le manifeste décrit la manière dont le complément est intégré dans les clients Outlook. Actuellement, il existe trois versions du schéma de manifeste, dont  **VersionOverrides**. Pour créer votre complément, nous vous recommandons d’utiliser la version 1.1 **VersionOverrides** 1.0 du schéma de manifeste. Voici un exemple.

 >**Remarque**  Dans l’exemple suivant, toutes les valeurs d’URL commencent par « YOUR_WEB_SERVER ». Cette valeur est un espace réservé. Dans un manifeste valide réelle, ces valeurs contiendraient des URL web HTTPS valides.




```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
        <Host Name="Mailbox" />
    </Hosts>
    <Requirements>
        <Sets>
            <Set Name="MailBox" MinVersion="1.1" />
        </Sets>
    </Requirements>
    <!-- These elements support older clients that don't support add-in commands -->
    <FormSettings>
        <Form xsi:type="ItemRead">
            <DesktopSettings>
                <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
                     on the ribbon in clients that support add-in commands. You can
                     use a completely different page if desired -->
                <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html" />
                <RequestedHeight>450</RequestedHeight>
            </DesktopSettings>
        </Form>
        <Form xsi:type="ItemEdit">
            <DesktopSettings>
                <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html" />
            </DesktopSettings>
        </Form>
    </FormSettings>
    <Permissions>ReadWriteItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    </Rule>
    <DisableEntityHighlighting>false</DisableEntityHighlighting>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

        <Requirements>
            <bt:Sets DefaultMinVersion="1.3">
                <bt:Set Name="Mailbox" />
            </bt:Sets>
        </Requirements>
        <Hosts>
            <Host xsi:type="MailHost">

                <DesktopFormFactor>
                    <FunctionFile resid="functionFile" />

                    <!-- Custom pane, only applies to read form -->
                    <ExtensionPoint xsi:type="CustomPane">
                        <RequestedHeight>100</RequestedHeight>
                        <SourceLocation resid="customPaneUrl" />
                        <Rule xsi:type="RuleCollection" Mode="Or">
                            <Rule xsi:type="ItemIs" ItemType="Message" />
                            <Rule xsi:type="ItemIs" ItemType="AppointmentAttendee" />
                        </Rule>
                    </ExtensionPoint>

                    <!-- Message compose form -->
                    <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="msgComposeDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="msgComposeFunctionButton">
                                    <Label resid="funcComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcComposeSuperTipTitle" />
                                        <Description resid="funcComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>addDefaultMsgToBody</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="msgComposeMenuButton">
                                    <Label resid="menuComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuComposeSuperTipTitle" />
                                        <Description resid="menuComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="msgComposeMenuItem1">
                                            <Label resid="menuItem1ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ComposeLabel" />
                                                <Description resid="menuItem1ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg1ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgComposeMenuItem2">
                                            <Label resid="menuItem2ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ComposeLabel" />
                                                <Description resid="menuItem2ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg2ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgComposeMenuItem3">
                                            <Label resid="menuItem3ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ComposeLabel" />
                                                <Description resid="menuItem3ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg3ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                                    <Label resid="paneComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneComposeSuperTipTitle" />
                                        <Description resid="paneComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="composeTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Appointment compose form -->
                    <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="apptComposeDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="apptComposeFunctionButton">
                                    <Label resid="funcComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcComposeSuperTipTitle" />
                                        <Description resid="funcComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>addDefaultMsgToBody</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="apptComposeMenuButton">
                                    <Label resid="menuComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuComposeSuperTipTitle" />
                                        <Description resid="menuComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="apptComposeMenuItem1">
                                            <Label resid="menuItem1ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ComposeLabel" />
                                                <Description resid="menuItem1ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg1ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptComposeMenuItem2">
                                            <Label resid="menuItem2ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ComposeLabel" />
                                                <Description resid="menuItem2ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg2ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptComposeMenuItem3">
                                            <Label resid="menuItem3ComposeLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ComposeLabel" />
                                                <Description resid="menuItem3ComposeTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>addMsg3ToBody</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                                    <Label resid="paneComposeButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneComposeSuperTipTitle" />
                                        <Description resid="paneComposeSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="composeTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Message read form -->
                    <ExtensionPoint xsi:type="MessageReadCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="msgReadDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="msgReadFunctionButton">
                                    <Label resid="funcReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcReadSuperTipTitle" />
                                        <Description resid="funcReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>getSubject</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="msgReadMenuButton">
                                    <Label resid="menuReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuReadSuperTipTitle" />
                                        <Description resid="menuReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="msgReadMenuItem1">
                                            <Label resid="menuItem1ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ReadLabel" />
                                                <Description resid="menuItem1ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemClass</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgReadMenuItem2">
                                            <Label resid="menuItem2ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ReadLabel" />
                                                <Description resid="menuItem2ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getDateTimeCreated</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="msgReadMenuItem3">
                                            <Label resid="menuItem3ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ReadLabel" />
                                                <Description resid="menuItem3ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemID</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                    <Label resid="paneReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneReadSuperTipTitle" />
                                        <Description resid="paneReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="readTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>

                    <!-- Appointment read form -->
                    <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                        <OfficeTab id="TabDefault">
                            <Group id="apptReadDemoGroup">
                                <Label resid="groupLabel" />
                                <!-- Function (UI-less) button -->
                                <Control xsi:type="Button" id="apptReadFunctionButton">
                                    <Label resid="funcReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="funcReadSuperTipTitle" />
                                        <Description resid="funcReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="blue-icon-16" />
                                        <bt:Image size="32" resid="blue-icon-32" />
                                        <bt:Image size="80" resid="blue-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>getSubject</FunctionName>
                                    </Action>
                                </Control>
                                <!-- Menu (dropdown) button -->
                                <Control xsi:type="Menu" id="apptReadMenuButton">
                                    <Label resid="menuReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="menuReadSuperTipTitle" />
                                        <Description resid="menuReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="red-icon-16" />
                                        <bt:Image size="32" resid="red-icon-32" />
                                        <bt:Image size="80" resid="red-icon-80" />
                                    </Icon>
                                    <Items>
                                        <Item id="apptReadMenuItem1">
                                            <Label resid="menuItem1ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem1ReadLabel" />
                                                <Description resid="menuItem1ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemClass</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptReadMenuItem2">
                                            <Label resid="menuItem2ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem2ReadLabel" />
                                                <Description resid="menuItem2ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getDateTimeCreated</FunctionName>
                                            </Action>
                                        </Item>
                                        <Item id="apptReadMenuItem3">
                                            <Label resid="menuItem3ReadLabel" />
                                            <Supertip>
                                                <Title resid="menuItem3ReadLabel" />
                                                <Description resid="menuItem3ReadTip" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="red-icon-16" />
                                                <bt:Image size="32" resid="red-icon-32" />
                                                <bt:Image size="80" resid="red-icon-80" />
                                            </Icon>
                                            <Action xsi:type="ExecuteFunction">
                                                <FunctionName>getItemID</FunctionName>
                                            </Action>
                                        </Item>
                                    </Items>
                                </Control>
                                <!-- Task pane button -->
                                <Control xsi:type="Button" id="apptReadOpenPaneButton">
                                    <Label resid="paneReadButtonLabel" />
                                    <Supertip>
                                        <Title resid="paneReadSuperTipTitle" />
                                        <Description resid="paneReadSuperTipDescription" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="green-icon-16" />
                                        <bt:Image size="32" resid="green-icon-32" />
                                        <bt:Image size="80" resid="green-icon-80" />
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <SourceLocation resid="readTaskPaneUrl" />
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>
                </DesktopFormFactor>
            </Host>
        </Hosts>

        <Resources>
            <bt:Images>
                <!-- Blue icon -->
                <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png" />
                <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png" />
                <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
                <!-- Red icon -->
                <bt:Image id="red-icon-16" DefaultValue="YOUR_WEB_SERVER/images/red-16.png" />
                <bt:Image id="red-icon-32" DefaultValue="YOUR_WEB_SERVER/images/red-32.png" />
                <bt:Image id="red-icon-80" DefaultValue="YOUR_WEB_SERVER/images/red-80.png" />
                <!-- Green icon -->
                <bt:Image id="green-icon-16" DefaultValue="YOUR_WEB_SERVER/images/green-16.png" />
                <bt:Image id="green-icon-32" DefaultValue="YOUR_WEB_SERVER/images/green-32.png" />
                <bt:Image id="green-icon-80" DefaultValue="YOUR_WEB_SERVER/images/green-80.png" />
            </bt:Images>
            <bt:Urls>
                <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html" />
                <bt:Url id="readTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html" />
                <bt:Url id="composeTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppCompose/TaskPane/TaskPane.html" />
                <bt:Url id="customPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/CustomPane/CustomPane.html" />
            </bt:Urls>
            <bt:ShortStrings>
                <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
                <!-- Compose mode -->
                <bt:String id="funcComposeButtonLabel" DefaultValue="Insert default message" />
                <bt:String id="menuComposeButtonLabel" DefaultValue="Insert message" />
                <bt:String id="paneComposeButtonLabel" DefaultValue="Insert custom message" />

                <bt:String id="funcComposeSuperTipTitle" DefaultValue="Inserts the default message" />
                <bt:String id="menuComposeSuperTipTitle" DefaultValue="Choose a message to insert" />
                <bt:String id="paneComposeSuperTipTitle" DefaultValue="Enter your own text to insert" />

                <bt:String id="menuItem1ComposeLabel" DefaultValue="Hello World!" />
                <bt:String id="menuItem2ComposeLabel" DefaultValue="Add-in commands are cool!" />
                <bt:String id="menuItem3ComposeLabel" DefaultValue="Visit dev.outlook.com" />

                <!-- Read mode -->
                <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
                <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
                <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

                <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
                <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
                <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

                <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
                <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
                <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
            </bt:ShortStrings>
            <bt:LongStrings>
                <!-- Compose mode -->
                <bt:String id="funcComposeSuperTipDescription" DefaultValue="Inserts text into body of the message or appointment. This is an example of a function button." />
                <bt:String id="menuComposeSuperTipDescription" DefaultValue="Inserts your choice of text into body of the message or appointment. This is an example of a drop-down menu button." />
                <bt:String id="paneComposeSuperTipDescription" DefaultValue="Opens a pane where you can enter text to insert in the body of the message or appointment. This is an example of a button that opens a task pane." />

                <bt:String id="menuItem1ComposeTip" DefaultValue="Inserts Hello World! into the body of the message or appointment." />
                <bt:String id="menuItem2ComposeTip" DefaultValue="Inserts Add-in commands are cool! into the body of the message or appointment." />
                <bt:String id="menuItem3ComposeTip" DefaultValue="Inserts Visit dev.outlook.com into the body of the message or appointment." />

                <!-- Read mode -->
                <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
                <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
                <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

                <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
                <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
                <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
            </bt:LongStrings>
        </Resources>
    </VersionOverrides>
</OfficeApp>

```


## Versions de schéma

Tous les clients Outlook ne prennent pas en charge les fonctionnalités les plus récentes en une seule fois et certains utilisateurs Outlook disposent d’une version antérieure d’Outlook. Les versions de schéma permettent aux développeurs de créer des compléments garantissant une compatibilité descendante, qui utilisent les fonctionnalités les plus récentes lorsqu’elles sont disponibles, mais fonctionnent avec des versions plus anciennes.

L’élément  **VersionOverrides** dans le manifeste en est un exemple. Tous les éléments définis dans **VersionOverrides** remplaceront le même élément dans l’autre partie du manifeste. Cela signifie que, dès que possible, Outlook utilisera les éléments de la section **VersionOverrides** pour configurer le complément. Toutefois, si la version d’Outlook ne prend pas en charge une version de **VersionOverrides**, Outlook l’ignorera et se référera aux informations contenues dans le reste du manifeste. 

Cette approche signifie que les développeurs ne doivent pas créer plusieurs manifestes individuels, mais plutôt conserver tous les éléments définis dans un fichier.

Les versions actuelles du schéma sont les suivantes :


|||
|:-----|:-----|
|Version|Description|
|v1.0|Prend en charge la version 1.0 de l’API JavaScript pour Office. Pour les compléments Outlook, la prise en charge des formulaires de lecture est également incluse. |
|v1.1|Prend en charge la version 1.1 de l’API JavaScript pour Office et  **VersionOverrides**. Pour les compléments Outlook, la prise en charge des formulaires de composition est incluse.|
|**VersionOverrides** 1.0|Prend en charge les versions ultérieures de l’API JavaScript pour Office. La prise en charge des commandes de complément est incluse.|
Cet article porte sur les conditions requises pour la version 1.1 du manifeste. Même si le manifeste de votre complément utilise l’élément  **VersionOverrides**, il est important d’inclure les éléments de la version 1.1 du manifeste afin que votre application fonctionne avec des clients plus anciens qui ne prennent pas en charge  **VersionOverrides**.


## Élément racine

L’élément racine du manifeste de complément Outlook est  **OfficeApp**. Cet élément indique également l’espace de noms, la version de schéma et le type de complément par défaut. Placez tous les autres éléments du manifeste entre ses balises d’ouverture et de fermeture. Vous trouverez ci-dessous un exemple d’élément racine :


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest>

</OfficeApp>
```


## Version

Il s’agit de la version du complément spécifique. Si un développeur met à jour un élément dans le manifeste, la version doit être incrémentée également. Ainsi, lorsque le nouveau manifeste est installé, il remplace celui existant et l’utilisateur a accès aux nouvelles fonctionnalités. Si ce complément a déjà été envoyé à l’Office Store, le nouveau manifeste devra à nouveau être soumis et validé. Ensuite, les utilisateurs de ce complément auront accès au nouveau manifeste mis à jour automatiquement en quelques heures, après son approbation.

If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in.


## VersionOverrides

L’élément **VersionOverrides** représente l’emplacement des informations pour les commandes de complément. Pour obtenir des informations supplémentaires sur cet élément, voir [Définir des commandes de complément dans votre manifeste de complément Outlook](../../outlook/manifests/define-add-in-commands.md).


## Localisation

Certains aspects du complément doivent être localisés pour les différents paramètres régionaux, tels que le nom, la description et l’URL qui est chargée. Ces éléments peuvent être facilement localisés en spécifiant la valeur par défaut et les valeurs de remplacement locales dans l’élément **Resources** au sein de l’élément **VersionOverrides**. Pour remplacer une image, une URL et une chaîne, procédez comme suit :


```XML
<Resources>
    <bt:Images>
      <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
        <!-- add information for other locales -->

    <bt:Urls>
      <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
        <!-- add information for other locales -->

    <bt:ShortStrings> 
      </bt:String>
      <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
        <bt:Override Locale="ar-sa" Value="<add localized value here>" />
        <!-- add information for other locales -->
    </bt:ShortStrings>

  </Resources>
```

La référence de schéma contient des informations complètes sur les éléments pouvant être localisés.


## Hôtes

Les compléments Outlook spécifient l’élément  **Hosts** comme indiqué ci-dessous.


```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Il existe une différence avec l’élément  **Hosts** au sein de l’élément **VersionOverrides**, qui fait l’objet de l’article [Définir des commandes de complément dans votre manifeste de complément Outlook](../../outlook/manifests/define-add-in-commands.md).


## Configuration requise

L’élément  **Requirements** spécifie l’ensemble d’API disponible pour le complément. Pour un complément Outlook, l’ensemble de conditions requises doit être Mailbox et avoir la valeur 1.1 ou supérieure. Reportez-vous à la référence d’API pour connaître la dernière version de condition requise. Pour plus d’informations sur les ensembles de conditions requises, voir [API de complément Outlook](../../outlook/apis.md).

L’élément  **Requirements** peut également apparaître dans l’élément **VersionOverrides**, ce qui permet au complément de spécifier d’autres conditions requises lorsqu’il est chargé dans des clients qui prennent en charge  **VersionOverrides**.

L’exemple suivant utilise l’attribut  **DefaultMinVersion** de l’élément **Sets** pour exiger office.js version 1.1 ou ultérieure, et l’attribut **MinVersion** de l’élément **Set** pour exiger l’ensemble de conditions requises Mailbox version 1.1.




```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```


## Paramètres de formulaire

L’élément  **FormSettings** est utilisé par les clients Outlook plus anciens, qui prennent en charge uniquement le schéma version 1.1 et non **VersionOverrides**. À l’aide de cet élément, les développeurs définissent la façon dont le complément s’affiche dans ces clients. Il existe deux parties : **ItemRead** et **ItemEdit**.  **ItemRead** est utilisé pour spécifier la manière dont le complément apparaît lorsque l’utilisateur lit les messages et les rendez-vous. **ItemEdit** décrit comment le complément s’affiche lorsque l’utilisateur compose une réponse, un nouveau message, un nouveau rendez-vous ou modifie un rendez-vous dont il est l’organisateur.

Ces paramètres sont directement liés aux règles d’activation dans l’élément  **Rule**. Par exemple, si un complément spécifie qu’il doit apparaître sur un message lors de sa composition, un formulaire  **ItemEdit** doit être spécifié.

Pour plus d’informations, voir [Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)](../../overview/add-in-manifests.md)

## Domaines d’application

Le domaine de la page de démarrage du complément que vous spécifiez dans l’élément  **SourceLocation** est le domaine par défaut pour le complément. Si vous n’utilisez pas les éléments **AppDomains** et **AppDomain** et que votre complément tente d’accéder à un autre domaine, le navigateur ouvre une nouvelle fenêtre en dehors du panneau de complément. Afin que le complément puisse accéder à un autre domaine dans le volet de complément, ajoutez un élément **AppDomains** et incluez chaque domaine supplémentaire dans son propre sous-élément **AppDomain** dans le manifeste de complément.

L’exemple suivant spécifie le domaine  `https://www.contoso2.com` comme second domaine auquel le complément peut accéder à l’intérieur du volet du complément :




```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

Les domaines d’application sont également nécessaires pour activer le partage entre la fenêtre contextuelle et le complément en cours d’exécution dans le client riche.


## Autorisations

L’élément  **Permissions** contient les autorisations requises pour le complément. Généralement, vous devez spécifier l’autorisation nécessaire minimale dont votre complément a besoin selon la méthode exacte que vous prévoyez d’utiliser. Par exemple, un complément de messagerie qui s’active dans les formulaires de composition et qui lit uniquement mais n’écrit pas dans les propriétés de l’élément comme [item.requiredAttendees](../../../reference/outlook/Office.context.mailbox.item.md), et qui n’appelle pas [mailbox.makeEwsRequestAsync](../../../reference/outlook/Office.context.mailbox.md) pour accéder aux opérations liées aux services web doit spécifier l’autorisation **ReadItem**. Pour plus de détails sur les autorisations disponibles, voir [Spécifier les autorisations pour l’accès du complément Outlook à la boîte aux lettres de l’utilisateur](../../outlook/understanding-outlook-add-in-permissions.md).


**Modèle d’autorisations à 4 niveaux pour les compléments de messagerie**

![Modèle d’autorisations à 4 niveaux pour le schéma d’applications de messagerie v1.1](../../../images/olowa15wecon_Permissions_4Tier.png)
```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>

```


## Règles d’activation

Les règles d’activation sont spécifiées dans l’élément  **Rule**. L’élément  **Rule** peut apparaître en tant qu’enfant de l’élément **OfficeApp** dans les manifestes version 1.1, ainsi qu’en tant qu’enfant de l’élément **ExtensionPoint** dans **VersionOverrides**. Voir [Définir des commandes de complément dans votre manifeste de complément Outlook](../../outlook/manifests/define-add-in-commands.md) pour plus de détails sur l’utilisation de cet élément dans **VersionOverrides**.

Les règles d’activation peuvent être utilisées pour activer un complément basé sur une ou plusieurs des conditions suivantes sur l’élément sélectionné.


- Le type d’élément et/ou la classe de message
    
- La présence d’un type spécifique d’entité connue, comme une adresse ou un numéro de téléphone
    
- Une correspondance d’expression régulière dans le corps, l’objet ou l’adresse e-mail de l’expéditeur
    
- L’existence d’une pièce jointe
    
Pour plus de détails et pour des exemples de règles d’activation, voir [Règles d’activation pour les compléments Outlook](../../outlook/manifests/activation-rules.md).


## Prochaines étapes : commandes de complément


Après avoir défini un manifeste de base, [définissez des commandes pour votre complément](../../outlook/manifests/define-add-in-commands.md). Les commandes de complément se présentent sous forme de bouton dans le ruban. Ainsi, les utilisateurs peuvent activer votre complément de façon simple et intuitive. Pour plus d’informations, voir [Commandes de complément pour Outlook](../../outlook/add-in-commands-for-outlook.md).


## Ressources supplémentaires



- [Compléments Outlook](../../outlook/outlook-add-ins.md)
    
- [Règles d’activation pour les compléments Outlook](../../outlook/manifests/activation-rules.md)
    
- [Localisation des compléments Office](../../develop/localization.md)
    
- [Créer un complément de messagerie pour Outlook qui s’exécute sur des ordinateurs de bureau, des tablettes et des appareils mobiles (schéma version 1.1)](http://msdn.microsoft.com/library/8d425fb3-8a7c-429d-87b3-8046e964b153%28Office.15%29.aspx)
    
- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../../outlook/privacy-and-security.md)
    
- [API de complément Outlook](../../outlook/apis.md)
    
- [Manifeste XML des compléments Office](../../overview/add-in-manifests.md)
    
- [Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)](../../overview/add-in-manifests.md)
    
- [Types d’éléments et classes de messages](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
    
- [Instructions de conception pour les compléments Office](../../design/add-in-design.md)
    
- [Présentation des autorisations de complément Outlook](../../outlook/understanding-outlook-add-in-permissions.md)
    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
    

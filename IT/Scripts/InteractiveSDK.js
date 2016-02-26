//Declare namespace
var InteractiveTutorial = {};
var Globals = {};

InteractiveTutorial.App = new function () {
    var _codeXml = null;
    var _contentXml = null;
    var _currentNodeId = null;
    var _currentTask = null;
    var _currentTaskIndex = null;
    var _currentContentIndex = null;
    var _currentScenario = null;
    var _currentScenarioIndex = null;
    var _tasks = null;
    var _checked = null;
    var _currentLink = null;
    var _contentList = null;
    var _editor = null;
    var self = this;

    this.init = function InteractiveTutorial_App$init() {
        _checked = {};
        _contentList = [];

        //Populates the _contentList from tutorial.xml file.
        self.getTutorials();

        $('#toastMessage .closeBtn').click(function () {
            $("#message").empty();
            $('#toastMessage').hide();
        });


        $("#run").click(self.executeCode);

        /*
        Escape key will move focus to the run code button.  
        Elements with a role of button will also react to enter and space keydown events like a input button
        */
        $("body").keydown(function (event) {
            if (event.which == 27) {
                $("#run").focus();
                $('#toastMessage').hide();
            } //Check for enter and space to mimic the behavior of buttons
            else if (event.which == 13 || event.which == 32) {
                var element = $(event.srcElement);
                if (element.attr("role") == "button") {
                    $(event.srcElement).click();
                }
            }
        });
    }

    //Shows tutorial list.
    this.showList = function InteractiveTutorial_App$showList() {
        $("#content").empty();
        $("#headercontent").empty();
        $("#content").attr("class", "listPageContent");
        $("#headercontent").append("<h1>Select a Tutorial</h1>");
        $("#headercontent").attr("class", "listHeader office-contentAccent1-color");
        $("#navigation").hide();
        var i;
        $("#content").append("<ul id='scenarioList'></ul>");
        var list = $("#scenarioList");
        for (i = 0; i < _contentList.length; i++) {
            var scenario = _contentList[i].scenario;
            var listItem = $("<li class='listItem office-contentAccent2-bgColor' role='button' tabindex='0'><div class='listText'><h2>" + self.htmlEncode(scenario) + "</h2><img src='Images/checkwhite.png' height='15px' alt='Check' /></div></li>");
            listItem.appendTo(list).click({ "content": _contentList, "index": i }, self.showAPIPage);
            if (!(_checked[scenario])) {
                listItem.find('img').hide();
            }
        }
    }

    //Execute code in text area. 
    this.executeCode = function InteractiveTutorial_App$executeCode() {
        try {
            // Execute the code found in the code textarea
            eval(_editor.getValue());
        }
        catch (err) {
            // Catch syntax and runtime errors
            showMessage('Error executing code: ' + err);
        }

    }

    //Gets the code snippet based on an id from code.xml. Uses third party library for code formatting and styling.
    this.getCodeSection = function InteractiveTutorial_App$getCodeSection(nodeId) {

        var getCodeCallback = function (xml) {
            var code = _codeXml.find('[id="' + nodeId + '"]').text();

            $('#codeWindow').val($.trim(code));
            _editor = CodeMirror.fromTextArea(document.getElementById('codeWindow'), {
                mode: "javascript",
                lineNumbers: false,
                matchBrackets: true
            });
            self.sizeCodeEditor();
        }

        //Improve performance, only get code.xml once, use stored version for other calls
        if (_codeXml == null) {
            self.getXml('code.xml', function (xml) {
                _codeXml = xml;
                $(window).resize(function () {
                    self.sizeCodeEditor();
                });
                getCodeCallback(xml);
            });
        }
        else {
            getCodeCallback(_codeXml);
        }
    }

    //Gets the content from tutorial.xml and adds to an array.
    this.getTutorials = function InteractiveTutorial_App$getTutorials() {
        self.getXml('tutorials.xml', function (xml) {
            _contentXml = xml;
            var tutorial = _contentXml.find('scenario').each(function () {
                var scenarioObject = {};
                var title = $(this).attr("title");
                var link = $(this).attr("link");
                var tasks = [];
                $(this).find("tasks").find("task").each(function () {
                    var taskObject = {};
                    var taskTitle = $(this).attr("title");
                    var taskId = $(this).attr("id");
                    var taskDescription = $(this).attr("description");
                    taskObject = { "title": taskTitle, "id": taskId, "description": taskDescription };
                    tasks.push(taskObject);
                });
                scenarioObject = { "scenario": title, "link": link, "tasks": tasks };
                _contentList.push(scenarioObject);
            });
            var scenario = _contentList[0].scenario;
            self.showList();
        });
    }

    //Gets xml from file.
    this.getXml = function InteractiveTutorial_App$getXml(url, callback) {
        // Get xml for the code and content
        $.ajax({
            url: url,
            cache: false,
            dataType: 'xml',
            success: function (xml) {
                callback($(xml));
            }
        });
    }

    //Loads current tutorial and associated next steps.
    this.showAPIPage = function InteractiveTutorial_App$showAPIPage(event) {
        $("#headercontent").attr("class", "whiteText office-contentAccent1-bgColor");
        $("#content").attr("class", "apiPageContent");
        _currentContentIndex = event.data.index;
        _currentScenario = event.data.content[_currentContentIndex].scenario;
        _currentLink = event.data.content[_currentContentIndex].link;
        _tasks = event.data.content[_currentContentIndex].tasks;
        _currentTaskIndex = 0;
        $("#headercontent").html("<div id='scenario'><span id='scenarioimg'><img src='Images/backwhite.png' role='button' tabindex='0' title='Back to Tutorial List' height='30px' alt='Back' /></span><span><div id='scenariolabel'><h4>" + self.htmlEncode(_currentScenario) + "</h4></div><div id='task'></div></span></div>").show();
        $("#scenarioimg img").click(self.showList);
        self.showTask();
    }

    //Show current step code and description.
    this.showTask = function InteractiveTutorial_App$showTask() {

        $("#content").html("");
        var APILayout = $("<div id='APILayout'></div>").appendTo("#content");
        var navigation = $("#navigation").html("").show();
        var isLastStep = (_currentTaskIndex == _tasks.length - 1);

        var taskTitle = _tasks[_currentTaskIndex].title;
        $("#task").html("<h3>" + self.htmlEncode(taskTitle) + "</h3>");
        var taskId = _tasks[_currentTaskIndex].id;
        var taskDescription = _tasks[_currentTaskIndex].description;

        if (taskId != "allCSSClass") {
            var menu = $("<div id='tabs'><ul'><li id='codeMenu' class='tabSelected office-contentAccent1-color'><a href='#' tabindex='0' title='View the code window'>CODE</a></li><li id='descriptionMenu'><a href='#' tabindex='0' title='View the description window'>DESCRIPTION</a></li></ul>").appendTo(APILayout);
            var codeMenu = $("#codeMenu").click(self.showCode);
            var descriptionMenu = $("#descriptionMenu").click(self.showDescription);
            var code = $("<div id='codeLayout'><textarea id='codeWindow' spellcheck='false'></textarea></div>").appendTo(APILayout);
            self.getCodeSection(taskId);
            var description = $("<div id='description'>" + self.htmlEncode(taskDescription, true) + "</div>").hide().appendTo(APILayout);

            var run = $("<button id='run' class='buttonclass' accesskey='R'><u>R</u>un Code</button>").appendTo(navigation).click(self.executeCode);
            $("#content").removeClass("cssPage");
        }
        else {
            $("#content").addClass("cssPage").html(self.getCSSClassesHTML());
        }

        $("<div class='navigationButtons'><div id='previous' role='button' title='Go to the previous step'></div><div id='next' role='button' tabindex='0' title='Go to the next step'><img src='Images/next.png' alt='Next' /></div></div>").appendTo(navigation);
        if (!isLastStep) {
            $("#next").click(function () {
                _currentTaskIndex++;
                self.showTask();
            });

            $("#next").mouseover(function () {
                $(this).find("img").attr("src", "Images/nextStep_hover.png");
            }).mouseout(function () {
                $(this).find("img").attr("src", "Images/next.png");
            });
        }
        else {
            $("#next").click(function () {
                _checked[_currentScenario] = true;
                self.showList();
            });

            $("#next").attr("title", "Show all tutorials").mouseover(function () {
                $(this).find("img").attr("src", "Images/list_hover.png");
            }).mouseout(function () {
                $(this).find("img").attr("src", "Images/list.png");
            });
            $("#next").find("img").attr("src", "Images/list.png");
        }

        if (_currentTaskIndex != 0) {
            $("<img src='Images/previous.png' alt='Previous' /></div>").appendTo("#previous");
            $("#previous").click(function () {
                _currentTaskIndex--;
                self.showTask();

            }).attr("tabindex", "0");

            $("#previous").mouseover(function () {
                $(this).find("img").attr("src", "Images/previous_hover.png");
            }).mouseout(function () {
                $(this).find("img").attr("src", "Images/previousStep.png");

            });
        }
    }

    this.getCSSClassesHTML = function () {
        var documentTheme =
            [
                { name: "office-docTheme-primary-fontColor" },
                { name: "office-docTheme-primary-bgColor" },
                { name: "office-docTheme-secondary-fontColor" },
                { name: "office-docTheme-secondary-bgColor" },
                { name: "office-contentAccent1-color" },
                { name: "office-contentAccent2-color" },
                { name: "office-contentAccent3-color" },
                { name: "office-contentAccent4-color" },
                { name: "office-contentAccent5-color" },
                { name: "office-contentAccent6-color" },
                { name: "office-contentAccent1-bgColor" },
                { name: "office-contentAccent2-bgColor" },
                { name: "office-contentAccent3-bgColor" },
                { name: "office-contentAccent4-bgColor" },
                { name: "office-contentAccent5-bgColor" },
                { name: "office-contentAccent6-bgColor" },
                { name: "office-contentAccent1-borderColor", isBorder: true },
                { name: "office-contentAccent2-borderColor", isBorder: true },
                { name: "office-contentAccent3-borderColor", isBorder: true },
                { name: "office-contentAccent4-borderColor", isBorder: true },
                { name: "office-contentAccent5-borderColor", isBorder: true },
                { name: "office-contentAccent6-borderColor", isBorder: true },
                { name: "office-headerFont-eastAsian" },
                { name: "office-headerFont-latin" },
                { name: "office-headerFont-script" },
                { name: "office-headerFont-localized" },
                { name: "office-bodyFont-eastAsian" },
                { name: "office-bodyFont-latin" },
                { name: "office-bodyFont-script" },
                { name: "office-bodyFont-localized" },
                { name: "office-body-bgColor" },
                { name: "office-officeTheme-primary-fontColor" },
                { name: "office-officeTheme-primary-bgColor" },
                { name: "office-officeTheme-secondary-fontColor" },
                { name: "office-officeTheme-secondary-bgColor" },
            ];

        var returnHtml = "";
        for (var i = 0; i < documentTheme.length; i++) {
            if (documentTheme[i].isBorder) {
                returnHtml += "<div style='border-style:solid;border-width:3px;' class='" + documentTheme[i].name + "'>" + documentTheme[i].name + "</div>";
            }
            else {
                returnHtml += "<div class='" + documentTheme[i].name + "'>" + documentTheme[i].name + "</div>";
            }
        }
        return returnHtml;
    }

    //Shows the code area tab.
    this.showCode = function InteractiveTutorial_App$showCode() {
        $("#codeMenu").attr("class", "tabSelected office-contentAccent1-color");
        $("#descriptionMenu").attr("class", "");
        $("#action").show();
        $("#description").hide();
        $("#codeLayout").show();

    }

    //Shows the description tab.
    this.showDescription = function InteractiveTutorial_App$showDescription() {
        $("#codeMenu").attr("class", "");
        $("#descriptionMenu").attr("class", "tabSelected office-contentAccent1-color");
        $("#action").hide();
        $("#codeLayout").hide();
        $("#description").show()

    }

    //Resizing code editor.
    this.sizeCodeEditor = function InteractiveTutorial_App$sizeCodeEditor() {
        $(".CodeMirror-scroll, #description").css("height", $("#content").height() - $("#tabs").height() - 10 + "px");
    }

    this.htmlEncode = function InteractiveTutorial_App$htmlEncode(value, allowLinks) {
        var allowedTags = ["<br />", "<br>", "<i>", "</i>", "<b>", "</b>"];
        var links = [];

        var encodedHTML = $("<div/>").html(value);
        if (allowLinks == true) {
            encodedHTML.find("a[href]").replaceWith(function (index, element) {
                var linkText = this.innerText;
                if (linkText == undefined) {
                    linkText = this.text;
                }
                links.push({ href: this.href, target: this.target, text: linkText, title: this.title });
                return "LINKPLACEHOLDER" + index;
            });
        }

        if (encodedHTML.html) {
            encodedHTML = encodedHTML.html();
        }

        encodedHTML = $("<div/>").text(encodedHTML).html();

        $.each(allowedTags, function (index, tag) {
            var regex = new RegExp($('<div/>').text(tag).html(), "gi");
            encodedHTML = encodedHTML.replace(regex, tag);
        }
         );

        $.each(links, function (index, element) {
            encodedHTML = encodedHTML.replace("LINKPLACEHOLDER" + index, "<a href='" + element.href + "' title='" + element.title + "' target='" + element.target + "'>" + element.text + "</a>");
        });

        return encodedHTML;
    }
}

// Display a message at the bottom of the task pane - called when code is executed in the window
function showMessage(text) {
    if (text != "" && text != null) {
        $("#message").html(InteractiveTutorial.App.htmlEncode(text, false) + "<br />" + $("#message").html());
    }
    $("#toastMessage").slideDown("fast");
}

Office.initialize = function (reason) {
/*
        var themeChange = new OfficeThemeManager();
        themeChange.onThemeChange(function (e) {
            showMessage(JSON.stringify(e.theme));
        });
        
        Office.context.document.getActiveViewAsync(function (asyncResult) {
            if (asyncResult.status == "succeeded") {
                var manageVisibility = new DetectVisibilityState(reason);
                manageVisibility.onVisibilityChange(function (e) {
                    showMessage(e.visible.toString());
                });
            }
        });
        var handler = function (args) {
        }

        Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, handler, function (asyncResult) {
        });
*/
        InteractiveTutorial.App.init();
}

////$(window).bind("resize", function () {
////    debugger;
////    showMessage("RESIZE");
////});

var DetectVisibilityState = function (reason) {
    var self = this;

    self.onAppSlide = false;

    var OM = Office.context.document;

    var appSlideData = { slideId: null, confidenceLevel: 0 };

    var ConfidenceLevel = {
        Low: 0,
        Medium: 1,
        High: 2
    };

    var currentView = "edit";

    var updateSlideData = function (slideId, confidenceLevel) {
        if (appSlideData.confidenceLevel >= confidenceLevel) {
            return;
        }
        appSlideData.slideId = slideId;
        appSlideData.confidenceLevel = confidenceLevel;
        OM.settings.set("appSlideData", appSlideData);
        OM.settings.saveAsync();
    }

    var getAndUpdateSlideData = function (confidenceLevel) {
        OM.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
            if (asyncResult.status === "failed") {
                throw asyncResult.error.message;
            }
            else {
                if (asyncResult.value.slides.count > 1) {
                    confidenceLevel = ConfidenceLevel.Low;
                }
                updateSlideData(asyncResult.value.slides[0].id, confidenceLevel);
            }
        });
    }

    $(window).bind("click keydown", function () {
        var confidenceLevel = ConfidenceLevel.High;
        if (confidenceLevel > appSlideData.confidenceLevel) {
            getAndUpdateSlideData(confidenceLevel);
        }
    });

    self.setActiveView = function (view) {
        currentView = view;
    };

    self.onVisibilityChange = function (callback) {
        var insideAgave = false;
        var currentStateSameSlide = false;
        var initialSlideCheckPerformed = false;

        //need to poll the host to see what the current slide is as no event is exposed for changing slides
        $(window).bind("mousemove focus", function () {
            insideAgave = true;
        });
        $(window).bind("mouseleave blur", function () {
            insideAgave = false;
        });
        window.setInterval(function () {
            if (currentView.toUpperCase() === "READ") {
                if ((!insideAgave || !currentStateSameSlide || !initialSlideCheckPerformed)) {
                    OM.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
                        if (asyncResult.status == "succeeded") {
                            //compare the current slide to what the app thinks its slide is, if different fire event
                            currentStateSameSlide = (asyncResult.value.slides[0].id == appSlideData.slideId);
                            if (self.onAppSlide != currentStateSameSlide || !initialSlideCheckPerformed) {
                                initialSlideCheckPerformed = true;
                                self.onAppSlide = currentStateSameSlide;
                                callback({ visible: currentStateSameSlide });
                            }
                        }
                    });
                }
            }
            else {
                initialSlideCheckPerformed = false;
            }
        }, 2000);
    };

    $(function () {
        var storedData = OM.settings.get("appSlideData");
        if (storedData != null) {
            appSlideData = storedData;
        }
        if (reason === "inserted") {
            getAndUpdateSlideData(ConfidenceLevel.Medium);
        }
        else {
            //handle the copy/paste scenario where "existing" gets returned, and the app already has a confidence level set
            OM.getActiveViewAsync(function (asyncResult) {
                //can't trust previous data if the app was copied
                if (asyncResult.value === "edit") {
                    appSlideData.confidenceLevel = ConfidenceLevel.Low;
                    getAndUpdateSlideData(ConfidenceLevel.Medium);
                }
                else {
                    //there are additional scenarios due to an existing app activated while user is on a different slide
                    //this line gives the user a chance to correct any mistakes by clicking in the app
                    appSlideData.confidenceLevel = ConfidenceLevel.Medium;
                }
            });

        }
    });
}
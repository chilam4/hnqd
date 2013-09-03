﻿var lastFrameTabId = 1;
var currentFrameTabId = 1;
var frameTabCount = 1;
var PE_FrameTab = {
    AddNew: function() {
        jQuery("#newFrameTab").click();
    }
    ,
    CloseCurrentTab: function(){
        jQuery("#iFrameTab" + currentFrameTabId).find(".closeTab").click();
    }
};

jQuery.fn.iFrameTab = function() {
    jQuery(this).each(function() {
        var cr = jQuery(this);
        var tabId = cr.attr("id").replace("iFrameTab", "");
        cr.click(function() {
            SwitchIframe(this);
        }
        ).find(".closeTab").click(function() {
            if (frameTabCount > 1) {
                var mainRightFrame = jQuery("#main_right_frame iframe[tabid='" + tabId + "']");
                var bClose = mainRightFrame[0].contentWindow.OnCloseTab ? mainRightFrame[0].contentWindow.OnCloseTab() : true;
                if (bClose) {
                    if (cr.attr("class") == "current") {
                        var nextIframe = cr.prev("li[id^='iFrameTab']");
                        if (nextIframe.length <= 0) {
                            nextIframe = cr.next("li[id^='iFrameTab']");
                        }
                        SwitchIframe(nextIframe[0]);
                    }

                    cr.remove();
                    jQuery("#frmTitle iframe[tabid='" + tabId + "']").remove();
                    mainRightFrame.remove();
                    frameTabCount--;
                    CheckFramesScroll();
                }
            }
        }
        ).end().dblclick(function() {
            jQuery(this).find(".closeTab").click();
        }
        );
    }
    );
    return jQuery(this);
}

function SwitchIframe(iFrameTab) {
    var tabId = jQuery(iFrameTab).attr("id").replace("iFrameTab", "");
    if (currentFrameTabId == tabId) {
        return false;
    }
    var switchFunc = jQuery("#main_right")[0].contentWindow.window.BeforeSwitch;
    var bSwitch = (switchFunc) ? switchFunc() : true;
    if (!bSwitch) {
        return false;
    }
    var currentGuideSrc = jQuery("#frmTitle iframe[tabid='" + currentFrameTabId + "']").attr("src");
    SetCurrentFrameTab(iFrameTab);
    var guideFrames = jQuery("#frmTitle > iframe").hide().attr({
        "id": "", "name": "" }
    );
    var mainFrames = jQuery("#main_right_frame > iframe").hide().attr({
        "id": "", "name": "" }
    );
    var newGuideFrame = jQuery("#frmTitle iframe[tabid='" + tabId + "']");
    var newMainFrame = jQuery("#main_right_frame iframe[tabid='" + tabId + "']");

    mainFrames.each(function() {
        this.contentWindow.window.name = "";
    }
    );
    guideFrames.each(function() {
        this.contentWindow.window.name = "";
    }
    );
    if (newGuideFrame.length <= 0) {
        newGuideFrame = jQuery("#frmTitle").append(jQuery("#iframeGuideTemplate").html())
        .find("[tabid=0]").attr({
            "tabid": tabId, "src": currentGuideSrc || "about:blank", "id": "left", "name": "left" }
        )
        .css("display", "block");
    }
    else {
        newGuideFrame = jQuery("#frmTitle iframe[tabid='" + tabId + "']")
        .attr("id", "left").attr("name", "left").show();
    }
    newGuideFrame[0].contentWindow.window.name = "left";
    frames["left"] = newGuideFrame[0].contentWindow.window;
    if (newMainFrame.length <= 0) {
        newMainFrame = jQuery("#main_right_frame").prepend(jQuery("#iframeMainTemplate").html())
        .find("[tabid=0]").attr({
            "tabid": tabId, "src": "about:blank", "id": "main_right", "name": "main_right" }
        )
        .css("display", "block");
    }
    else {
        newMainFrame = jQuery("#main_right_frame iframe[tabid='" + tabId + "']")
        .attr("id", "main_right").attr("name", "main_right").show();
    }
    newMainFrame[0].contentWindow.window.name = "main_right";
    frames["main_right"] = newMainFrame[0].contentWindow.window;
    currentFrameTabId = tabId;
    resizeFrame();
    var switchInto = jQuery("#main_right")[0].contentWindow.window.SwitchInto;
    if(switchInto){
        switchInto();
    }
}

function InitNewFrameTab() {
    jQuery("#newFrameTab").click(function() {

        jQuery('<li id="iFrameTab' + (++lastFrameTabId) + '" ><a href="javascript:"><span id="frameTabTitle">(无标题)</span><a class="closeTab"><img border="0" src="' + StyleSheetPath + 'images/tab-close.gif"/></a></a></li>')
        .insertBefore(this).iFrameTab();
        frameTabCount++;
        SwitchIframe(jQuery("#iFrameTab" + lastFrameTabId)[0]);
        if (CheckFramesScroll()) {
            jQuery("#FrameTabs ul:eq(0)").css("margin-left", cW - fW - 40);
        }
    }
    );
}

function NewFrameTab() {
    jQuery("#newFrameTab").click();
}

function SetCurrentFrameTab(selector) {
    jQuery("#FrameTabs .current").removeClass("current");
    jQuery(selector).addClass("current");
}

function CheckFramesScroll() {
    var ft = jQuery("#FrameTabs");
    window.cW = ft.width();
    window.fW = ft.find("ul:eq(0)").width();
    ft.unbind("DOMMouseScroll").unbind("mousewheel");
    if (fW > cW) {
        if (jQuery.browser.mozilla) {
            ft.bind("DOMMouseScroll", function(e) {
                ScrollFrames(cW, fW, e);
            }
            );
        }
        else {
            ft.bind("mousewheel", function(e) {
                ScrollFrames(cW, fW, e);
            }
            );
        }
        jQuery("#FrameTabs .tab-strip-wrap").addClass("tab-strip-margin");
        jQuery("#FrameTabs .tab-right, #FrameTabs .tab-left").css("display", "block");
        return true;
    }
    else {
        jQuery("#FrameTabs ul:eq(0)").css("margin-left", 0);
        jQuery("#FrameTabs .tab-right, #FrameTabs .tab-left").css("display", "none");
        jQuery("#FrameTabs .tab-strip-wrap").removeClass("tab-strip-margin");
        return false;
    }
}

function ScrollFrames(cW, fW, event, y) {
    if (!y) {
        if (event.wheelDelta) {
            y = event.wheelDelta / 5;
        }
        else if (event.detail) {
            y = -event.detail * 8;
        }
    }
    var jList = jQuery("#FrameTabs ul:eq(0)");
    var ml = jList.css("margin-left");
    ml = Number(ml.toLowerCase().replace("px", ""));
    if ((ml < 0 && y > 0) || (ml - cW > -fW - 40) && y < 0) {
        ml = ml + y;
        if (ml >= 0) {
            ml = 0;
        }
        if (ml - cW <= -fW - 40) {
            ml = cW - fW - 40;
        }
        jList.css("margin-left", ml);
    }
}

function RegScrollFramesBtn() {
    jQuery("#FrameTabs .tab-right").click(function() {
        ScrollFrames(window.cW,window.fW,null,-50);
    }
    );
    jQuery("#FrameTabs .tab-left").click(function() {
        ScrollFrames(window.cW,window.fW,null,50);
    }
    );
}

function SetTabTitle(tarFrame) {
    var title = "";
    try {
        title = tarFrame.contentWindow.document.title;
    }
    catch (e) {
    }
    var subTitle = title = title || "(无标题)";
    if (title.length > 6) {
        subTitle = title.substr(0, 5) + ".." }
    jQuery("#iFrameTab" + jQuery(tarFrame).attr("tabid")).find("#frameTabTitle").html(subTitle).attr("title", title);
}

function resizeFrame() {
    var width = document.body.clientWidth - 207;
    var lHeight = document.body.clientHeight - 78;
    var rHeight = lHeight - (jQuery("#FrameTabs").height() || 0) ;
    document.getElementById("main_right").style.width = width > 0 ? width : 0;
    document.getElementById("main_right").style.height = rHeight > 0 ? rHeight : 0;
    document.getElementById("left").style.height = lHeight > 0 ? lHeight : 0;
    jQuery("#FrameTabs").width(width);
}

jQuery(function() {
    jQuery("#FrameTabs li[id^='iFrameTab']").iFrameTab();
    InitNewFrameTab();
    RegScrollFramesBtn();
}
);
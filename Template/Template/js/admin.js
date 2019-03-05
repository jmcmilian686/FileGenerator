if (typeof jQuery === "undefined") {
    throw new Error("jQuery plugins need to be before this file");
}

$.AdminBSB = {};
$.AdminBSB.options = {
    colors: {
        red: '#F44336',
        pink: '#E91E63',
        purple: '#9C27B0',
        deepPurple: '#673AB7',
        indigo: '#3F51B5',
        blue: '#2196F3',
        lightBlue: '#03A9F4',
        cyan: '#00BCD4',
        teal: '#009688',
        green: '#4CAF50',
        lightGreen: '#8BC34A',
        lime: '#CDDC39',
        yellow: '#ffe821',
        amber: '#FFC107',
        orange: '#FF9800',
        deepOrange: '#FF5722',
        brown: '#795548',
        grey: '#9E9E9E',
        blueGrey: '#607D8B',
        black: '#000000',
        white: '#ffffff'
    },
    leftSideBar: {
        scrollColor: 'rgba(0,0,0,0.5)',
        scrollWidth: '4px',
        scrollAlwaysVisible: false,
        scrollBorderRadius: '0',
        scrollRailBorderRadius: '0',
        scrollActiveItemWhenPageLoad: true,
        breakpointWidth: 1170
    },
    dropdownMenu: {
        effectIn: 'fadeIn',
        effectOut: 'fadeOut'
    }
}

/* Left Sidebar - Function =================================================================================================
*  You can manage the left sidebar menu options
*  
*/
$.AdminBSB.leftSideBar = {
    activate: function () {
        var _this = this;
        var $body = $('body');
        var $overlay = $('.overlay');

        //Close sidebar
        $(window).click(function (e) {
            var $target = $(e.target);
            if (e.target.nodeName.toLowerCase() === 'i') { $target = $(e.target).parent(); }

            if (!$target.hasClass('bars') && _this.isOpen() && $target.parents('#leftsidebar').length === 0) {
                if (!$target.hasClass('js-right-sidebar')) $overlay.fadeOut();
                $body.removeClass('overlay-open');
            }
        });

        $.each($('.menu-toggle.toggled'), function (i, val) {
            $(val).next().slideToggle(0);
        });

        //When page load
        $.each($('.menu .list li.active'), function (i, val) {
            var $activeAnchors = $(val).find('a:eq(0)');

            $activeAnchors.addClass('toggled');
            $activeAnchors.next().show();
        });

        //Collapse or Expand Menu
        $('.menu-toggle').on('click', function (e) {
            var $this = $(this);
            var $content = $this.next();

            if ($($this.parents('ul')[0]).hasClass('list')) {
                var $not = $(e.target).hasClass('menu-toggle') ? e.target : $(e.target).parents('.menu-toggle');

                $.each($('.menu-toggle.toggled').not($not).next(), function (i, val) {
                    if ($(val).is(':visible')) {
                        $(val).prev().toggleClass('toggled');
                        $(val).slideUp();
                    }
                });
            }

            $this.toggleClass('toggled');
            $content.slideToggle(320);
        });

        //Set menu height
        _this.setMenuHeight();
        _this.checkStatuForResize(true);
        $(window).resize(function () {
            _this.setMenuHeight();
            _this.checkStatuForResize(false);
        });

        //Set Waves
        Waves.attach('.menu .list a', ['waves-block']);
        Waves.init();
    },
    setMenuHeight: function (isFirstTime) {
        if (typeof $.fn.slimScroll != 'undefined') {
            var configs = $.AdminBSB.options.leftSideBar;
            var height = ($(window).height() - ($('.legal').outerHeight() + $('.user-info').outerHeight() + $('.navbar').innerHeight()));
            var $el = $('.list');

            $el.slimscroll({
                height: height + "px",
                color: configs.scrollColor,
                size: configs.scrollWidth,
                alwaysVisible: configs.scrollAlwaysVisible,
                borderRadius: configs.scrollBorderRadius,
                railBorderRadius: configs.scrollRailBorderRadius
            });

            //Scroll active menu item when page load, if option set = true
            if ($.AdminBSB.options.leftSideBar.scrollActiveItemWhenPageLoad) {
                var activeItemOffsetTop = $('.menu .list li.active')[0].offsetTop
                if (activeItemOffsetTop > 150) $el.slimscroll({ scrollTo: activeItemOffsetTop + 'px' });
            }
        }
    },
    checkStatuForResize: function (firstTime) {
        var $body = $('body');
        var $openCloseBar = $('.navbar .navbar-header .bars');
        var width = $body.width();

        if (firstTime) {
            $body.find('.content, .sidebar').addClass('no-animate').delay(1000).queue(function () {
                $(this).removeClass('no-animate').dequeue();
            });
        }

        if (width < $.AdminBSB.options.leftSideBar.breakpointWidth) {
            $body.addClass('ls-closed');
            $openCloseBar.fadeIn();
        }
        else {
            $body.removeClass('ls-closed');
            $openCloseBar.fadeOut();
        }
    },
    isOpen: function () {
        return $('body').hasClass('overlay-open');
    }
};
//==========================================================================================================================

/* Right Sidebar - Function ================================================================================================
*  You can manage the right sidebar menu options
*  
*/
$.AdminBSB.rightSideBar = {
    activate: function () {
        var _this = this;
        var $sidebar = $('#rightsidebar');
        var $overlay = $('.overlay');

        //Close sidebar
        $(window).click(function (e) {
            var $target = $(e.target);
            if (e.target.nodeName.toLowerCase() === 'i') { $target = $(e.target).parent(); }

            if (!$target.hasClass('js-right-sidebar') && _this.isOpen() && $target.parents('#rightsidebar').length === 0) {
                if (!$target.hasClass('bars')) $overlay.fadeOut();
                $sidebar.removeClass('open');
            }
        });

        $('.js-right-sidebar').on('click', function () {
            $sidebar.toggleClass('open');
            if (_this.isOpen()) { $overlay.fadeIn(); } else { $overlay.fadeOut(); }
        });
    },
    isOpen: function () {
        return $('.right-sidebar').hasClass('open');
    }
}
//==========================================================================================================================

/* Searchbar - Function ================================================================================================
*  You can manage the search bar
*  
*/
var $searchBar = $('.search-bar');
$.AdminBSB.search = {
    activate: function () {
        var _this = this;

        //Search button click event
        $('.js-search').on('click', function () {
            _this.showSearchBar();
        });

        //Close search click event
        $searchBar.find('.close-search').on('click', function () {
            _this.hideSearchBar();
        });

        //ESC key on pressed
        $searchBar.find('input[type="text"]').on('keyup', function (e) {
            if (e.keyCode == 27) {
                _this.hideSearchBar();
            }
        });
    },
    showSearchBar: function () {
        $searchBar.addClass('open');
        $searchBar.find('input[type="text"]').focus();
    },
    hideSearchBar: function () {
        $searchBar.removeClass('open');
        $searchBar.find('input[type="text"]').val('');
    }
}
//==========================================================================================================================

/* Navbar - Function =======================================================================================================
*  You can manage the navbar
*  
*/
$.AdminBSB.navbar = {
    activate: function () {
        var $body = $('body');
        var $overlay = $('.overlay');

        //Open left sidebar panel
        $('.bars').on('click', function () {
            $body.toggleClass('overlay-open');
            if ($body.hasClass('overlay-open')) { $overlay.fadeIn(); } else { $overlay.fadeOut(); }
        });

        //Close collapse bar on click event
        $('.nav [data-close="true"]').on('click', function () {
            var isVisible = $('.navbar-toggle').is(':visible');
            var $navbarCollapse = $('.navbar-collapse');

            if (isVisible) {
                $navbarCollapse.slideUp(function () {
                    $navbarCollapse.removeClass('in').removeAttr('style');
                });
            }
        });
    }
}
//==========================================================================================================================

/* Input - Function ========================================================================================================
*  You can manage the inputs(also textareas) with name of class 'form-control'
*  
*/
$.AdminBSB.input = {
    activate: function () {
        //On focus event
        $('.form-control').focus(function () {
            $(this).parent().addClass('focused');
        });

        //On focusout event
        $('.form-control').focusout(function () {
            var $this = $(this);
            if ($this.parents('.form-group').hasClass('form-float')) {
                if ($this.val() == '') { $this.parents('.form-line').removeClass('focused'); }
            }
            else {
                $this.parents('.form-line').removeClass('focused');
            }
        });

        //On label click
        $('body').on('click', '.form-float .form-line .form-label', function () {
            $(this).parent().find('input').focus();
        });

        //Not blank form
        $('.form-control').each(function () {
            if ($(this).val() !== '') {
                $(this).parents('.form-line').addClass('focused');
            }
        });
    }
}
//==========================================================================================================================

/* Form - Select - Function ================================================================================================
*  You can manage the 'select' of form elements
*  
*/
$.AdminBSB.select = {
    activate: function () {
        if ($.fn.selectpicker) { $('select:not(.ms)').selectpicker(); }
    }
}
//==========================================================================================================================

/* DropdownMenu - Function =================================================================================================
*  You can manage the dropdown menu
*  
*/

$.AdminBSB.dropdownMenu = {
    activate: function () {
        var _this = this;

        $('.dropdown, .dropup, .btn-group').on({
            "show.bs.dropdown": function () {
                var dropdown = _this.dropdownEffect(this);
                _this.dropdownEffectStart(dropdown, dropdown.effectIn);
            },
            "shown.bs.dropdown": function () {
                var dropdown = _this.dropdownEffect(this);
                if (dropdown.effectIn && dropdown.effectOut) {
                    _this.dropdownEffectEnd(dropdown, function () { });
                }
            },
            "hide.bs.dropdown": function (e) {
                var dropdown = _this.dropdownEffect(this);
                if (dropdown.effectOut) {
                    e.preventDefault();
                    _this.dropdownEffectStart(dropdown, dropdown.effectOut);
                    _this.dropdownEffectEnd(dropdown, function () {
                        dropdown.dropdown.removeClass('open');
                    });
                }
            }
        });

        //Set Waves
        Waves.attach('.dropdown-menu li a', ['waves-block']);
        Waves.init();
    },
    dropdownEffect: function (target) {
        var effectIn = $.AdminBSB.options.dropdownMenu.effectIn, effectOut = $.AdminBSB.options.dropdownMenu.effectOut;
        var dropdown = $(target), dropdownMenu = $('.dropdown-menu', target);

        if (dropdown.length > 0) {
            var udEffectIn = dropdown.data('effect-in');
            var udEffectOut = dropdown.data('effect-out');
            if (udEffectIn !== undefined) { effectIn = udEffectIn; }
            if (udEffectOut !== undefined) { effectOut = udEffectOut; }
        }

        return {
            target: target,
            dropdown: dropdown,
            dropdownMenu: dropdownMenu,
            effectIn: effectIn,
            effectOut: effectOut
        };
    },
    dropdownEffectStart: function (data, effectToStart) {
        if (effectToStart) {
            data.dropdown.addClass('dropdown-animating');
            data.dropdownMenu.addClass('animated dropdown-animated');
            data.dropdownMenu.addClass(effectToStart);
        }
    },
    dropdownEffectEnd: function (data, callback) {
        var animationEnd = 'webkitAnimationEnd mozAnimationEnd MSAnimationEnd oanimationend animationend';
        data.dropdown.one(animationEnd, function () {
            data.dropdown.removeClass('dropdown-animating');
            data.dropdownMenu.removeClass('animated dropdown-animated');
            data.dropdownMenu.removeClass(data.effectIn);
            data.dropdownMenu.removeClass(data.effectOut);

            if (typeof callback == 'function') {
                callback();
            }
        });
    }
}
//==========================================================================================================================

/* Browser - Function ======================================================================================================
*  You can manage browser
*  
*/
var edge = 'Microsoft Edge';
var ie10 = 'Internet Explorer 10';
var ie11 = 'Internet Explorer 11';
var opera = 'Opera';
var firefox = 'Mozilla Firefox';
var chrome = 'Google Chrome';
var safari = 'Safari';

$.AdminBSB.browser = {
    activate: function () {
        var _this = this;
        var className = _this.getClassName();

        if (className !== '') $('html').addClass(_this.getClassName());
    },
    getBrowser: function () {
        var userAgent = navigator.userAgent.toLowerCase();

        if (/edge/i.test(userAgent)) {
            return edge;
        } else if (/rv:11/i.test(userAgent)) {
            return ie11;
        } else if (/msie 10/i.test(userAgent)) {
            return ie10;
        } else if (/opr/i.test(userAgent)) {
            return opera;
        } else if (/chrome/i.test(userAgent)) {
            return chrome;
        } else if (/firefox/i.test(userAgent)) {
            return firefox;
        } else if (!!navigator.userAgent.match(/Version\/[\d\.]+.*Safari/)) {
            return safari;
        }

        return undefined;
    },
    getClassName: function () {
        var browser = this.getBrowser();

        if (browser === edge) {
            return 'edge';
        } else if (browser === ie11) {
            return 'ie11';
        } else if (browser === ie10) {
            return 'ie10';
        } else if (browser === opera) {
            return 'opera';
        } else if (browser === chrome) {
            return 'chrome';
        } else if (browser === firefox) {
            return 'firefox';
        } else if (browser === safari) {
            return 'safari';
        } else {
            return '';
        }
    }
}
//==========================================================================================================================

$(function () {
    $.AdminBSB.browser.activate();
    $.AdminBSB.leftSideBar.activate();
    $.AdminBSB.rightSideBar.activate();
    $.AdminBSB.navbar.activate();
    $.AdminBSB.dropdownMenu.activate();
    $.AdminBSB.input.activate();
    $.AdminBSB.select.activate();
    $.AdminBSB.search.activate();

    setTimeout(function () { $('.page-loader-wrapper').fadeOut(); }, 50);
});



//================================Parse Table===========================================================================

function arrayify(collection) {
    return Array.prototype.slice.call(collection);
}

/**
 * generates factory functions to convert table rows to objects,
 * based on the titles in the table's <thead>
 * @param  {Array[String]} headings the values of the table's <thead>
 * @return {Function}      a function that takes a table row and spits out an object
 */
function factory(headings) {
    return function (row) {
        return arrayify(row.cells).reduce(function (prev, curr, i) {
            prev[headings[i]] = curr.innerText;
            return prev;
        }, {});
    }
}

/**
 * given a table, generate an array of objects.
 * each object corresponds to a row in the table.
 * each object's key/value pairs correspond to a column's heading and the row's value for that column
 * 
 * @param  {HTMLTableElement} table the table to convert
 * @return {Array[Object]}       array of objects representing each row in the table
 */
function parseTable(table) {
    var headings = arrayify(table.tHead.rows[0].cells).map(function (heading) {
        return heading.innerText;
    });
    return arrayify(table.tBodies[0].rows).map(factory(headings));
}



//======================Fields Handling Area========================================================================================


$("#btn_plus").click(function (e) {
    console.log("clicked");
    e.preventDefault();
    $('#largeModal').modal();
    
});

var idDel = 0;

$(".btnDel").click(function (e) {
    e.preventDefault;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            idDel = $(this).text();
        }

    })
    $("#defaultModal").modal();
    console.log(idDel);

});


$("#clModal").click(function (e) {
    e.preventDefault();
    console.log("clicked");
    
   /* $.get('Field/Details', function (data) {
        $('#tabDetail').html(data);
    }); */
});


$("#delFinal").click(function (e) {

    e.preventDefault();
    if (idDel > 0) {

        var url1 = '/Field/Delete/' + idDel;
        $.ajax({
            // edit to add steve's suggestion.
            //url: "/ControllerName/ActionName",
            url: url1,
            method: 'POST',  // post
            success: function (data) {
                if (data == "ok") {
                    $("#updArea").html("<h4> Field Deleted Successfully!</h4>");
                    $("#delFinal").hide();
                    $("#clModal").addClass("bg-green");
                    $("#clModal").attr('onclick', 'location.href = \'/Field/Index\'');
                }
              
            }
        });

    }

});

// Edit Action
$(".btnEdit").click(function (e) {
    console.log("clicked");
    var id = -1;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            id = $(this).text();
        }

    })
    var url1 = '/Field/Edit/' + id ;

    $.ajax({
        // edit to add steve's suggestion.
        //url: "/ControllerName/ActionName",
        url:  url1,
        success: function (data) {
           
            $("#largeModal").html(data);
        }
    });
    $('#largeModal').modal();
 

});

// Save Action
$("#largeModal").on("submit", "#form-fieldedit", function (e) {
    e.preventDefault();  // prevent standard form submission

    var form = $(this);
    $.ajax({
        url: form.attr("action"),
        method: form.attr("method"),  // post
        data: form.serialize(),
        success: function (partialResult) {
            $("#largeModal").html(partialResult);
        }
    });
});


//====================== End of Fields Handling Area========================================================================================

//======================Struct Handling Area========================================================================================
$("#btnPlus2").click(function (e) {
    console.log("clicked");
    e.preventDefault();
    $('#strModal').modal();

});


// Save Action
$("#strModal").on("submit", "#form-structedit", function (e) {
    e.preventDefault();  // prevent standard form submission

    var form = $(this);
    $.ajax({
        url: form.attr("action"),
        method: form.attr("method"),  // post
        data: form.serialize(),
        success: function (partialResult) {
            $("#strModal").html(partialResult);
        }
    });
});

// Edit Action
$(".btnEditS").click(function (e) {
    console.log("clicked");
    var id = -1;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            id = $(this).text();
        }

    })

    console.log(id);
    var url1 = '/Structure/Edit/' + id;

    $.ajax({
       
        url: url1,
        success: function (data) {

            $("#strModal").html(data);
        }
    });
    $('#strModal').modal();


});

//Delete Action


$(".btnDelS").click(function (e) {
    e.preventDefault;
    idDel = 0;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            idDel = $(this).text();
        }

    })
    $("#defaultModalS").modal();
    console.log(idDel);

});



$("#delFinalS").click(function (e) {

    e.preventDefault();
    if (idDel > 0) {

        var url1 = '/Structure/Delete/' + idDel;
        $.ajax({
            // edit to add steve's suggestion.
            //url: "/ControllerName/ActionName",
            url: url1,
            method: 'POST',  // post
            success: function (data) {
                if (data == "ok") {
                    $("#updAreaS").html("<h4> Field Deleted Successfully!</h4>");
                    $("#delFinalS").hide();
                    $("#clModalS").addClass("bg-green");
                    $("#clModalS").attr('onclick', 'location.href = \'/Structure/Index\'');
                }

            }
        });

    }

});
//======================End of Struct Handling Area========================================================================================

//======================StructFile Handling Area========================================================================================
$("#btnPlusSF").click(function (e) {
    console.log("clicked");
    e.preventDefault();
    $('#FSModal').modal();

});


// Save Action
$("#FSModal").on("submit", "#form-fielstruct", function (e) {
    e.preventDefault();  // prevent standard form submission

    var form = $(this);
    $.ajax({
        url: form.attr("action"),
        method: form.attr("method"),  // post
        data: form.serialize(),
        success: function (partialResult) {
            $("#FSModal").html(partialResult);
        }
    });
});


// Edit Action
$(".btnEditSF").click(function (e) {
    console.log("clicked");
    var id = -1;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            id = $(this).text();
        }

    })

    console.log(id);
    var url1 = '/Fieldstruct/Edit/' + id;

    $.ajax({

        url: url1,
        method:"post",
        success: function (data) {

            $("#FSModal").html(data);
        }
    });
    $('#FSModal').modal();


});


//Delete Action


$(".btnDelS").click(function (e) {
    e.preventDefault;
    idDel = 0;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            idDel = $(this).text();
        }

    })
    $("#defaultModalFS").modal();
    console.log(idDel);

});



$("#delFinalFS").click(function (e) {

    e.preventDefault();
    if (idDel > 0) {

        var url1 = '/Fieldstruct/Delete/' + idDel;
        $.ajax({
            // edit to add steve's suggestion.
            //url: "/ControllerName/ActionName",
            url: url1,
            method: 'POST',  // post
            success: function (data) {
                if (data == "ok") {
                    $("#updAreaFS").html("<h4> Field Deleted Successfully!</h4>");
                    $("#delFinalFS").hide();
                    $("#clModalFS").addClass("bg-green");
                    $("#clModalFS").attr('onclick', 'location.href = \'/Fieldstruct/Index\'');
                }

            }
        });

    }

});

//======================End of StructFile Handling Area========================================================================================

//======================Document Handling Area========================================================================================
$("#btnPlusDS").click(function (e) {
    console.log("clicked");
    e.preventDefault();
    $('#DSeModal').modal();

});

         //++++++++++++Document Structs handling+++++++++++++++

$(document).on("click","#stPlusBtn",function (e) {
    e.preventDefault();
    console.log("clicked");
    var table = document.querySelector("#strTable");

    var tbl = parseTable(table);
    /* $('table#strTable tr').map(function () {
        return $(this).find('td').not('.btns').map(function () {
            return $(this).html();
        }).get();
    }).get();*/

    var m

    var modelBck = {
        StructurePos: tbl,
        Posit: $("#position").val(),
        NewStId: $("#strId").val()
    }
    console.log(modelBck);

    $.ajax({
        url: '/Document/AddPos/',
        method: 'POST',
        success: function (data) {
            $('#structPos').html(data);
        },
        data: modelBck
    });
});

$(document).on("click", ".remTab", function (e) {
    e.preventDefault();
    console.log("clicked");
    console.log($(this).parents().closest('tr'));
    $(this).parents().closest('tr').remove();

});//+++++++++++++++++ End of Document Struct Handling+++++++++++++++++++++++++++++++++++++

// Save Action
$("#DSeModal").on("submit", "#form-Docedit", function (e) {
    e.preventDefault();  // prevent standard form submission
    //collect data
    var table = document.querySelector("#strTable");

    var tbl = parseTable(table);
    var structures = {
        StructurePos: tbl,
        Posit: 0,
        NewStId: 0
    };

    var FileID = $("#Document_LFile_ID").val();

    if (FileID == "") {
        FileID = 0;
    }

    var docObj = {
        LFile_ID: FileID,
        Doc_Name: $("#Doc_Name").val(),
        Doc_Description: $("#Doc_Desc").val(),
        Num_Struct: $("#Num_Struct").val()

    }
    var modelBck = {
        Document: docObj,
        DocStructs: structures

    };
    console.log(modelBck);

    var form = $(this);
    $.ajax({
        url: form.attr("action"),
        method: form.attr("method"),  // post
        data: modelBck,
        success: function (partialResult) {
            $("#DSeModal").html(partialResult);
        }
    });
});

// Edit Document-------------------

$(".btnEditSE").click(function (e) {
    console.log("clicked");
    var id = -1;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            id = $(this).text();
        }

    })

    console.log(id);
    var url1 = '/Document/Edit/' + id;

    $.ajax({

        url: url1,
        success: function (data) {

            $("#DSeModal").html(data);
        }
    });
    $('#DSeModal').modal();


});

//Delete Action


$(".btnDelSE").click(function (e) {
    e.preventDefault;
    idDel = 0;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            idDel = $(this).text();
        }

    })
    $("#defaultModalSE").modal();
    console.log(idDel);

});



$("#delFinalSE").click(function (e) {
    console.log(idDel);
    e.preventDefault();
    if (idDel > 0) {

        var url1 = '/Document/Delete/' + idDel;
        $.ajax({
            // edit to add steve's suggestion.
            //url: "/ControllerName/ActionName",
            url: url1,
            method: 'POST',  // post
            success: function (data) {
                if (data == "ok") {
                    $("#updAreaSE").html("<h4> Document Deleted Successfully!</h4>");
                    $("#delFinalSE").hide();
                    $("#clModalSE").addClass("bg-green");
                    $("#clModalSE").attr('onclick', 'location.href = \'/Document/Index\'');
                }

            }
        });

    }

});

$("#table1 tr").editable({

    keyboard: true,
    dblclick: true, 
    button: true

});

$("#addition").click(function (e) {
    e.preventDefault();

    var dict = {
        key:0,
        value:0
    };

    var id = $("#fldId").val();

    var mod = {
        ID: id,
        DataValue: dict

    };

    $.ajax({

        url: '/Datafield/Table/',
        data:mod,
        success: function (data) {

            $("#renderTable").html(data);
            $(".addData").each(function () {
                console.log($(this));
                $(this).prop('required',true);
            });
        }
    });


});

$("#linking").click(function (e) {
    e.preventDefault();

    
    var id = $("#fldId").val();

    var arr = [];
    $(".addData").each(function () {
        var Elements = {
            Name: $(this).attr('name'),
            Val: $(this).val()

        };

        arr.push(Elements);


    });

   

    var mod = {
        ID: id,
        DataValue: arr

    };
    console.log(mod);

    $.ajax({

        url: '/Datafield/Table/',
        data: mod,
        method: "post",
        success: function (data) {

            $("#renderTable").html(data);
            $(".addData").each(function () {
                console.log($(this));
                $(this).prop('required', true);
            });
        }
    });


});

$("#dataSave").on('click', function (e) {
    e.preventDefault();
   

    var arr = [];
    $(".addData").each(function () {
        var Elements = {
            Name: $(this).attr('name'),
            Val: $(this).val()

        };
       
       arr.push(Elements);
       
       
    });

    console.log(arr);

    var mod = {
        ID: 0,
        DataValue: arr

    };


    console.log(mod);

    $.ajax({

        url: '/Datafield/Create/',
        method:'post',
        data: mod,
        success: function (data) {
            if (data == "ok") {
                $(".form-insert").remove();
                $("#modalOk").modal();
            } else {
                $(".form-insert").remove();
                $("#modalCont").html("<p>" + data + "</p>");
                $("#modalOk").modal();
                

            }
        }
    });



});

$("#sel_file").on("click", function (e) {

    e.preventDefault();
    var n = $(".addData").length;
    console.log(n);

   

        $("#file1").click();

        console.log($("#file1").val());

   

});

$("#file1").change(function () {
    $('#selectedF').text($('#file1')[0].files[0].name);
});


$("#dataUpl").on("click", function (e) {

    e.preventDefault();

    var files = $('#file1')[0].files;


    

    if (files.length > 0) {

        if (window.FormData !== undefined) {

            var data = new FormData();
            for (var x = 0; x < files.length; x++) {
                data.append("file" + x, files[x]);
            }

            console.log(data);

            $.ajax({
                type: "POST",
                url: '/Datafield/UpFile' ,
                contentType: false,
                processData: false,
                data: data,
                beforeSend: function () {
                    $('.page-loader-wrapper').fadeIn();
                },
                success: function (result) {
                    console.log(result);
                    $('.page-loader-wrapper').fadeOut();
                },
                error: function (xhr, status, p3, p4) {
                    var err = "Error " + " " + status + " " + p3 + " " + p4;
                    if (xhr.responseText && xhr.responseText[0] == "{")
                        err = JSON.parse(xhr.responseText).Message;
                    console.log(err);
                    $('.page-loader-wrapper').fadeOut();
                }
            });

        }


    }

});



//+++++++++++ Linked Button Press+++++++++++++++++++++

$(".btnLinV").click(function(e) {
    e.preventDefault();
    var id = -1;
    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            id = $(this).text();
        }

    });

    console.log(id);

    $.ajax({

        url: '/Datafield/LinkedList/'+id,
        method: 'post',
        success: function (data) {
            $("#updArea").html(data);
            $("#delFinalData").hide();
            $("#defaultModal").modal();

        }
    });


});

var idDelete = 0;

$(".btnDelData").click(function (e) {

    e.preventDefault();
    idDelete = 0;

    $(this).parents().siblings().each(function () {

        if ($(this).hasClass("idElem")) {
            idDelete = $(this).text();
        }

    });

    var data = "<h4> Do you want to delete the selected Data?</h4>"
    $("#updArea").html(data);
    $("#delFinalData").show();
    $("#defaultModal").modal();


});


$("#delFinalData").click(function (e) {

    e.preventDefault();
       
    console.log(idDelete);

    $.ajax({

        url: '/Datafield/Delete/' + idDelete,
        method: 'post',
        success: function (data) {
            if (data == "ok") {

                var newdata = "<h4>Data was removed<h4>"
                $("#updArea").html(newdata);
            } else {

                 $("#updArea").html(data);
            }
           
            $("#delFinalData").hide();
            $("#defaultModal").modal();
            $("#clModalData").attr('onclick', 'location.href = \'/Datafield/Index\'');

        }
    });



});

//======================End of Document Handling Area========================================================================================

//=====================File Generation Area=================================================================================================

$("#customize").click(function (e) {

    e.preventDefault();

    $("#genFileModal").modal();

});

$("#addCustomize").click(function (e) {

    e.preventDefault();

    var doc = $("#docId").val();

    var field = $("#fieldId").val();
    var name = $("#fieldId").find('option:selected').text();
    console.log(doc);

    if ((doc !== "") && (field !=="")) {
       
        var mod = {
            Field: field
        };

        
        $.ajax({

            url: '/Genfile/Comboshow/',
            method: 'post',
            data: mod,
            success: function (data) {

                console.log(data);

                if (data.length > 0) {

                   

                    var listOpt = "";

                    $.each(data, function (i, item) {
                        console.log(item);
                        listOpt += '<option value="' + item.ID + '">' + item.Data + '</option>';
                    });

                    var newList = `<div class="form-group row grupp" style="margin-bottom:5px;">
                                   <div class="col-md-3">
                                    <label>`+ name + `:</label>
                                   </div> 
                                   <div class="col-md-5 clist">
                                       <select class="form-control col-md-8 show-tick valFilt" id="` + field + `">
                                                <option value="">-- Please select --</option>
                                                `+ listOpt + `
                                       </select>
                                   </div>
                                   <div class="col-md-1">
                                        <label>%:</label>
                                    </div>     
                                    <div class="col-md-2">
                                        
                                       <input type="number" class="form-control col-md-8 show-tick valPerc textbox" placeholder="%" value="100" max="100" min="0" style="border-bottom: 1px solid #ddd"/>
                                      
                                   </div>
                                   <div class="col-md-1">
                                       <button type="button" class="  btnDelFilt btn btn-table bg-deep-orange waves-effect" >
                                                <i class="material-icons">remove</i>
                                       </button>
                                   </div>
                                   </div>`;

                    $("#GenFilePar").append(newList);

                } else {
                    console.log(mod);
                    $.ajax({
                        url: '/Genfile/Nameshow/', method: 'post', data: mod, success: function (data) {

                            var newTextb = `<div class="form-group row grupp" style="margin-bottom:5px;">
                                   <div class="col-md-4">
                                    <label>`+ name + `</label>
                                   </div> 
                                   <div class="col-md-7">
                                       <input type="text" class="form-control col-md-8 show-tick valFilt textbox" id="` + data + `" style="border-bottom: 1px solid #ddd"/>
                                      
                                   </div>
                                  
                                   <div class="col-md-1">
                                       <button type="button" class="  btnDelFilt btn btn-table bg-deep-orange waves-effect">
                                                <i class="material-icons">remove</i>
                                       </button>
                                   </div>
                                   </div>`;

                            $("#GenFilePar").append(newTextb);



                        }
                    });

                    

                }

               
                    
            }
        });

    }

    

});


$(document).on("click", ".btnDelFilt", function (e) {

    e.preventDefault();
    $(this).parents().closest('.grupp').remove();;


});

$("#GenFileSave").click(function (e) {

    e.preventDefault();
    var arr = [];

    $(".valFilt").each(function () {

        if ($(this).hasClass("textbox")) {

            var Elements = {
                Name: $(this).attr("id"),
                Val: $(this).val(),
                Type: "text",
                Percent: 100
            };
            arr.push(Elements);

        } else {
            var Elements2 = {
                Name: $(this).find('option:selected').text(),
                Val: $(this).val(),
                Type: "list",
                Percent: $(this).parent().siblings().find('.valPerc').val()

            };
            arr.push(Elements2);
        }
        

       


    });

    var model = {

        DocID: $("#docId").val(),
        NDocs: $("#NumFiles").val(),
        NDets: $("#NumDetail").val(),
        NBatch: $("#NBatch").val(),
        FileName: $("#FileName").val(),
        Max:   $("#Multiple").val(),
        FiltValues: arr

    };


    console.log(model);
    var urls = $("#GenFileSave").attr("httpad");

    console.log(urls);

    $.ajax({

        url: urls,
        method: 'post',
        data: model,
        beforeSend: function () {
            $('.page-loader-wrapper').fadeIn();
        },
        success: function (data) {
            $("#genFileModal").html(data);
            $('.page-loader-wrapper').fadeOut();
         

        }
    });

});

//=====================End of File Generation Area==========================================================================================

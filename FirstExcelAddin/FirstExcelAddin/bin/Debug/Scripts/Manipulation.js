$(document).ready(function () {

    $('.selectedItemTab').css({ 'font-weight': 'bold' });
    $('.selectedItemTabHeader').css({ 'font-weight': 'bold' });
    $('.selectedItemTab').text($('.worksheetTab .active').text() + " details");
    $('.selectedItemTabHeader').text($('.worksheetTab .active').text() + " description");

    $('a[data-toggle="tab"]').on('shown.bs.tab', function () {
        var selectedTab = $('.worksheetTab .active').text() + " details";
        var selectedTabHeader = $('.worksheetTab .active').text() + " description";
        $('.selectedItemTab').text(selectedTab);
        $('.selectedItemTabHeader').text(selectedTabHeader);
    });
});



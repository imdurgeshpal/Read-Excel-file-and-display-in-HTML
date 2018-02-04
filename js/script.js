function handleFile(e) {
    //Get the files from Upload control
    var files = e.target.files;
    var i, f;
    var card_header_footer_color = $("#card_header_footer_color").val();
    var card_profile_site_color = $("#card_profile_site_color").val();
    var session_header = $("#session_header").val();
    var month_value = $('#month_value').val();
    var calendar_background = $('#calendar-background').val();

    //Loop through files

    for (i = 0, f = files[i]; i != files.length; ++i) {
        var reader = new FileReader();
        var name = f.name;
        reader.onload = function (e) {
            var data = e.target.result;

            var result;
            var workbook = XLSX.read(data, { type: 'binary' });

            var sheet_name_list = workbook.SheetNames;
            sheet_name_list.forEach(function (y) { /* iterate through sheets */
                //Convert the cell value to Json
                var roa = XLSX.utils.sheet_to_json(workbook.Sheets[y]);
                if (roa.length > 0) {
                    result = roa;
                }
            });

            //Get the first column first cell value
            for (i = 0; i < result.length; i++) {

                var x =
                    '<div class="col-md-4 mt-4 mb-4">' +
                    ' <div class="card">' +
                    ' <div class="card-header">' + result[i]['Training Title'] + '   </div>' +
                    ' <img class="card-img-top" src="' + result[i].Link + '">' +
                    ' <div class="card-block">' +
                    '  <figure class="profile">' +
                    '<span class="profile-avatar">' + result[i]['Date'] + ' </span>' +
                    ' </figure>' +
                    '<h4 class="card-title mt-3"> ' + result[i]['Training Objective'] + '  </h4>' +

                    '<div class="card-text"> Speaker: ' + result[i].Facilitator + '  </div>' +
                    ' </div>' +
                    '<div class="card-footer text-muted">' +
                    ' Training Duration : ' + result[i].Hours +
                    ' <span class="btn btn-info float-right">' + result[i].Location + '</span>' +
                    ' </div>' +
                    '</div>' +

                    ' </div>'
                    ;

                if (result[i]['Training Category'] == 'Open') {
                    $('#open-session').append(x);
                }
                else {
                    $('#close-session').append(x);
                }
                $(".card-header,.card-footer").css("background-color", card_header_footer_color);
                $(".profile").css("background-color", card_profile_site_color);
                $(".btn-info").css({ "background-color": card_profile_site_color, "border-color": card_profile_site_color });
                $(".open-session-header,.close-session-header").css("background-color", session_header);
                $('.overlay_file_chooser').hide(500);
                $('.month-text').html(month_value + "-2018");
                $('.container').css("background-color", calendar_background);
            }
        };
        reader.readAsArrayBuffer(f);
    }
}
function makeTable(array) {
    var table = document.createElement('table');
    for (var i = 0; i < array.length; i++) {
        var row = document.createElement('tr');
        for (var j = 0; j < array[i].length; j++) {
            var cell = document.createElement('td');
            cell.textContent = array[i][j];
            row.appendChild(cell);
        }
        table.appendChild(row);
    }
    return table;
}

//Change event to dropdownlist
$(document).ready(function () {
    $('#files').change(handleFile);
});
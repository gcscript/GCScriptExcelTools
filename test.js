function ExecuteScript() {
    var listContainer = document.querySelector('.jobs-search-results-list');

    var scrollStep = 200;
    var scrollInterval = 50;

    (function scrollList() {
        listContainer.scrollTop += scrollStep;

        // Stop scrolling when the bottom of the list is reached
        if (listContainer.scrollTop + listContainer.offsetHeight < listContainer.scrollHeight) {
            setTimeout(scrollList, scrollInterval);
        }
        else {
            listContainer.scrollTop = 0;
            ChangeColorJob();
        }
    })();
}

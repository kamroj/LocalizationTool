const $ = require('jquery');

module.exports = {
    showMessage: (message, warning) => {
        let div = document.getElementById('description');
        let text = document.createTextNode('~ ' + message + '\n');
        div.appendChild(text);

        if (warning) {
            let span = document.createElement('span');
            span.style = 'color:red;';
            span.appendChild(text);
            div.appendChild(span);
        }
    },

    blockButton: (ID, blocked) => {
        let div = document.getElementById(ID);
        if (blocked) {
            div.setAttribute('disabled', 'disabled');
            div.setAttribute('type', 'createButtonDisabled');
        } else {
            div.removeAttribute('disabled');
            div.setAttribute('type', 'createButtonEnabled');
        }
    },

    clearDiv: (divId) =>{        
        $(divId).empty();
    }
};

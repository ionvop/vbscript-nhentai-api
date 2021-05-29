class nhentaiAPI {
    constructor() {
        this.#resetVariables();
    }

    #resetVariables() {
        this.statusCode = this.pages = this.favorites = this.code = 0;
        this.title = this.uploaded = this.fullTitle = "";
        this.parodies = this.tags = this.artists = this.groups = this.languages = this.categories = [];
        this.exists = false;
    }

    #midString(input, left, right) {
        input = input.substring(input.indexOf(left) + left.length);
        input = input.substring(0, input.indexOf(right));
        return input;
    }

    #midStringList(input, left, right) {
        let list = [];

        while(input.includes(left) && input.includes(right)) {
            input = input.substring(input.indexOf(left) + left.length);
            list.push(input.substring(0, input.indexOf(right)));
            input = input.substring(input.indexOf(right) + right.length);
        }

        return list;
    }

    BookID(code) {
        if (this.code == code) {
            return this;
        }

        this.#resetVariables();
        this.code = code;
        let XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
        let xhr = new XMLHttpRequest();
        let data = "";
        xhr.open("GET", "https://nhentai.net/g/" + code + "/", false);
        xhr.send();
        data = xhr.responseText;
        this.statusCode = xhr.status;

        if (this.statusCode != 200) {
            if (this.statusCode == 404) {
                this.exists = false;
            }

            return this;
        }

        this.exists = true;
        this.title = this.#midString(data, "<meta itemprop=\"name\" content=\"", "\" />");
        this.parodies = this.#midString(data, "Parodies:", "Characters:");
        this.parodies = this.#midStringList(this.parodies, "<span class=\"name\">", "</span>");
        this.tags = this.#midString(data, "<meta name=\"twitter:description\" content=\"", "\" />");
        this.tags = this.tags.split(", ");
        this.artists = this.#midString(data, "Artists:", "Groups:");
        this.artists = this.#midStringList(this.artists, "<span class=\"name\">", "</span>");
        this.groups = this.#midString(data, "Groups:", "Languages:");
        this.groups = this.#midStringList(this.groups, "<span class=\"name\">", "</span>");
        this.languages = this.#midString(data, "Languages:", "Categories:");
        this.languages = this.#midStringList(this.languages, "<span class=\"name\">", "</span>");
        this.categories = this.#midString(data, "Categories:", "Pages:");
        this.categories = this.#midStringList(this.categories, "<span class=\"name\">", "</span>");
        this.pages = this.#midString(data, "Pages:", "Uploaded:");
        this.pages = this.#midString(this.pages, "<span class=\"name\">", "</span>");
        this.pages = Number(this.pages);
        this.uploaded = this.#midString(data, "Uploaded:", "ript>");
        this.uploaded = this.#midString(this.uploaded, "datetime", "<sc");
        this.uploaded = this.#midString(this.uploaded, ">", "<");
        this.favorites = this.#midString(data, "<i class=\"fas fa-heart\"></i><span>Favorite <span class=\"nobold\">(", ")");
        this.favorites = Number(this.favorites);
        this.fullTitle = this.#midString(data, "<h1 class=\"title\">", "</h1>");
        this.fullTitle = this.#midString(this.fullTitle, ">", "<");
        this.cover = this.#midString(data, "<meta itemprop=\"image\" content=\"", "\" />");
        return this;
    }
}

api = new nhentaiAPI;

console.log(api.BookID(215600).statusCode);
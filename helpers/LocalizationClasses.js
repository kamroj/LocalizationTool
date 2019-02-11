//Localizacja tekstów
class LocalizationContainer {
    constructor() {
        this.name;
        this.localizationPacks = [];
    }
}

class LocalizationPack {
    constructor() {
        this.name;
        this.localizationUnits = [];
    }
}

class LocalizationUnit {
    constructor() {
        this.name;
        this.localizationStrings = [];
    }
}

class LocalizationString {
    constructor() {
        this.key;
        this.value;
    }
}

//Localizacja aktorów
class ActorLocalizationPack {
    constructor() {
        this.actorLocalisationStrings = [];
    }
}

class ActorLocalizationString {
    constructor() {
        this.key;
        this.actorName;
        this.actorGender;
    }
}

module.exports = {
    LocalizationContainer: LocalizationContainer,
    LocalizationPack: LocalizationPack,
    LocalizationUnit: LocalizationUnit,
    LocalizationString: LocalizationString,
    ActorLocalizationPack: ActorLocalizationPack,
    ActorLocalizationString: ActorLocalizationString
};

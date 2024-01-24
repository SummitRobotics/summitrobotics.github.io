var config_data = `
{
  "dataFormat": "kvs",
  "enable_google_sheets": "true",
  "title": "Scouting PASS 2023",
  "page_title": "Charged Up",
  "checkboxAs": "10",
  "prematch": [
    { "name": "Scouter Initials",
      "code": "s",
      "gsCol": "s",
      "type": "scouter",
      "size": 5,
      "maxSize": 5,
      "required": "true"
    },
    { "name": "Event",
      "code": "e",
      "gsCol": "e",
      "type": "event",
      "defaultValue": "2023wayak",
      "required": "true",
      "disabled": "true"
    },
    { "name": "Match Level",
      "code": "l",
      "gsCol": "l",
      "type": "level",
      "choices": {
        "qm": "Quals<br>",
        "sf": "Semifinals<br>",
        "f": "Finals"
      },
      "defaultValue": "qm",
      "required": "true"
    },
    { "name": "Match #",
      "code": "m",
      "gsCol": "m",
      "type": "match",
      "min": 1,
      "max": 100,
      "required": "true"
    },
    { "name": "Robot",
      "code": "r",
      "gsCol": "r",
      "type": "robot",
      "choices": {
        "r1": "Red-1",
        "b1": "Blue-1<br>",
        "r2": "Red-2",
        "b2": "Blue-2<br>",
        "r3": "Red-3",
        "b3": "Blue-3"
      },
      "required":"true"
    },
    { "name": "Team #",
      "code": "t",
      "gsCol": "t",
      "type": "team",
      "min": 1,
      "max": 99999
    },
    { "name": "Auto Start Position",
      "code": "as",
      "gsCol": "as",
      "type": "clickable_image",
      "filename": "2023/field_image.png",
      "clickRestriction": "one",
      "allowableResponses": "2 3 4 9 10 11 14 15 16 21 22 23 26 27 28 33 34 35 38 39 46 47",
      "shape": "circle 5 black red true"
    }
  ],
  "auton": [
    { "name": "Auto Scoring",
      "code": "asg",
      "gsCol": "asg",
      "type": "clickable_image",
      "filename": "2023/grid_image.png",
      "dimensions": "9 4",
      "clickRestriction": "onePerBox",
      "toggleClick": "true",
      "showFlip": "true",
      "showUndo": "false",
      "shape": "circle 12 black red true"
    },
    { "name": "Game Pieces attempted<br>(Scored and Missed)",
      "code": "aa",
      "gsCol": "aa",
      "type": "counter"
    },
    { "name": "Mobility?",
      "code": "am",
      "gsCol": "am",
      "type": "bool"
    },
    { "name": "Docked",
      "code": "ad",
      "gsCol": "ad",
      "type":"radio",
      "choices": {
        "d": "Docked (not Engaged)<br>",
        "e": "Engaged<br>",
        "a": "Attempted but failed<br>",
        "x": "Not attempted"
      },
      "defaultValue": "x"
    }
  ],
  "teleop": [
    { "name": "Grid Scoring",
      "code": "tsg",
      "gsCol": "tsg",
      "type": "clickable_image",
      "filename": "2023/grid_image.png",
      "dimensions": "9 4",
      "clickRestriction": "onePerBox",
      "toggleClick": "true",
      "showFlip": "false",
      "showUndo": "false",
      "shape": "circle 12 black red true"
    },
    { "name": "Feeder Count<br>(Fed another bot)",
      "code": "tfc",
      "gsCol": "tfc",
      "type": "counter"
    },
    { "name": "Was Fed<br>Game Pieces",
      "code": "wf",
      "gsCol": "wf",
      "type": "bool"
    },
    { "name": "Was Defended",
      "code": "wd",
      "gsCol": "wd",
      "type": "bool"
    },
    { "name": "Who Defended this bot",
      "code": "who",
      "gsCol": "who",
      "type": "text"
    },
    { "name": "Smart Placement<br>(creates Links)",
      "code": "lnk",
      "gsCol": "lnk",
      "type": "bool"
    },
    { "name": "Floor Pickup",
      "code": "fpu",
      "gsCol": "fpu",
      "type": "radio",
      "choices": {
        "o": "Cones<br>",
        "u": "Cubes<br>",
        "b": "Both<br>",
        "x": "Not Attempted"
      },
      "defaultValue": "x"
    },
    { "name": "Substation Use",
      "code": "sub",
      "gsCol": "sub",
      "type": "radio",
      "choices": {
        "1": "Single<br>",
        "2": "Double<br>",
        "b": "Both<br>",
        "x": "Not Attempted"
      },
      "defaultValue": "x"
    }
  ],
  "endgame": [
    { "name": "Docking Timer",
      "code": "dt",
      "gsCol": "dt",
      "type": "timer"
    },
    { "name": "Final Status",
      "code": "fs",
      "gsCol": "fs",
      "type":"radio",
      "choices": {
        "p": "Parked<br>",
        "d": "Docked (Not Engaged)<br>",
        "e": "Engaged<br>",
        "a": "Attempted but failed<br>",
        "x": "Not attempted"
      },
      "defaultValue": "x"
    },
    { "name": "Total # of alliance<br>robots docked/engaged",
      "code": "dn",
      "gsCol": "dn",
      "type": "counter"
    },
    { "name": "Links Scored<br>(by alliance)",
      "code": "ls",
      "gsCol": "ls",
      "type": "counter"
    }
  ],
  "postmatch": [
    { "name": "Driver Skill",
      "code": "ds",
      "gsCol": "ds",
      "type": "radio",
      "choices": {
        "n": "Not Effective<br>",
        "a": "Average<br>",
        "v": "Very Effective<br>",
        "x": "Not Observed"
      },
      "defaultValue": "x"
    },
    { "name": "Defense Rating",
      "code": "dr",
      "gsCol": "dr",
      "type": "radio",
      "choices": {
        "b": "Below Average<br>",
        "a": "Average<br>",
        "g": "Good<br>",
        "e": "Excellent<br>",
        "x": "Did not play defense"
      },
      "defaultValue": "x"
    },
    { "name": "Speed Rating",
      "code": "sr",
      "gsCol": "sr",
      "type": "radio",
      "choices": {
        "1": "1 (slow)<br>",
        "2": "2<br>",
        "3": "3<br>",
        "4": "4<br>",
        "5": "5 (fast)"
      },
      "defaultValue":"3"
    },
    { "name": "Died/Immobilized",
      "code": "die",
      "gsCol": "die",
      "type": "bool"
    },
    { "name": "Tippy<br>(almost tipped over)",
      "code": "tip",
      "gsCol": "tip",
      "type": "bool"
    },
    { "name": "Dropped Cones (>2)",
      "code": "dc",
      "gsCol": "dc",
      "type": "bool"
    },
    { "name": "Make good<br>alliance partner?",
      "tooltip": "Would you want this robot on your alliance in eliminations?",
      "code": "all",
      "gsCol": "all",
      "type": "bool"
    },
    { "name": "Comments",
      "code": "co",
      "gsCol": "co",
      "type": "text",
      "size": 15,
      "maxSize": 55
    }
  ]
}`;

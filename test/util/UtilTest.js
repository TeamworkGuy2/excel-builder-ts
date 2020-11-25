"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var chai = require("chai");
var Util = require("../../util/Util");
var asr = chai.assert;
suite("Util", function UtilTest() {
    var defaultSettings = {
        hardErrors: true,
        maxLimit: 100,
        strictMath: true,
        plugins: ["default"],
    };
    test("defaults", function defaultsTest() {
        asr.deepEqual(Util.defaults({ name: "shark", strictMath: false }, {}), { name: "shark", strictMath: false });
        asr.deepEqual(Util.defaults({ name: "bob", maxLimit: 50, defaultNamespace: "blue" }, defaultSettings), { name: "bob", hardErrors: true, maxLimit: 50, strictMath: true, defaultNamespace: "blue", plugins: ["default"] /*, position: undefined */ });
    });
    test("pick", function pickTest() {
        asr.deepEqual(Util.pick(defaultSettings, ["defaultNamespace", "hardErrors"]), { /*defaultNamespace: undefined, */ hardErrors: true });
        asr.deepEqual(Util.pick({}, []), {});
        asr.deepEqual(Util.pick({}, ["blue", "shark"]), {});
    });
    test("uniqueId", function uniqueIdTest() {
        asr.equal(Util.uniqueId("") - Util.uniqueId(""), -1);
        asr.equal(Util.uniqueId("excel-builder-ts-qDVo8Z3QDW"), 1);
        asr.equal(Util.uniqueId("excel-builder-ts-qDVo8Z3QDW"), 2);
    });
    test("positionToLetterRef", function positionToLetterRefTest() {
        asr.equal(Util.positionToLetterRef(0, 0), "0");
        asr.equal(Util.positionToLetterRef(1, 3), "A3");
        asr.equal(Util.positionToLetterRef(34, 155), "AH155");
        asr.equal(Util.positionToLetterRef(26, 1), "Z1");
        asr.equal(Util.positionToLetterRef(34, 12), "AH12");
        asr.equal(Util.positionToLetterRef(676, 2), "YZ2");
        asr.equal(Util.positionToLetterRef(731, 35), "ABC35");
    });
});

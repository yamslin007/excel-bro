import {expect, it} from "vitest";
import {sameJsonEvidence} from "./jsonEvidence";
it("ignores only object key order", () => {
 expect(sameJsonEvidence({a:1,b:[{x:2,y:3}]},{b:[{y:3,x:2}],a:1})).toBe(true);
 for (const other of [{a:"1",b:[{x:2,y:3}]},{a:1,b:[{x:3,y:2}]},{a:1,b:[{x:2,y:3}],extra:null}])
   expect(sameJsonEvidence({a:1,b:[{x:2,y:3}]},other)).toBe(false);
 expect(sameJsonEvidence([1,2],[2,1])).toBe(false);
 expect(sameJsonEvidence(null,{})).toBe(false);
});

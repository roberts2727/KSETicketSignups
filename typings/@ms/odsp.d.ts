// Type definitions for Microsoft ODSP projects
// Project: ODSP

/* Global definition for UNIT_TEST builds
   Code that is wrapped inside an if(UNIT_TEST) {...}
   block will not be included in the final bundle when the
   --ship flag is specified */
declare const UNIT_TEST: boolean;
interface ArrayConstructor {
    from<T, U>(arrayLike: ArrayLike<T>, mapfn: (v: T, k: number) => U, thisArg?: any): Array<U>;
    from<T>(arrayLike: ArrayLike<T>): Array<T>;
}
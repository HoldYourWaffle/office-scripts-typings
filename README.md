# office-scripts-typings
TypeScript type definitions for Office Scripts.

## Generator
The Office Scripts runtime environment is a little _odd_, so we need to generate some type definitions.
You need two sets of typings:

1. [`excel.d.ts`](https://github.com/OfficeDev/office-scripts-docs-reference/blob/main/generate-docs/script-inputs/excel.d.ts), which contains the types for the Office Scripts API. Unfortunately this file [is not conveniently available](https://github.com/OfficeDev/office-scripts-docs-reference/issues/304#issuecomment-1834628886).
2. `misc.d.ts`, which contains some miscellaneous type definitions that are missing from the standard ES lib files, such as `console.log`.

The [`office-scripts-typings`](/src/office-scripts-typings.ts) script will take care of this for you.
Run the following command:

```sh
npx office-scripts-typings generate [output]
```

The generated type definitions will be put in the `output` directory. The default value is `@types/office-scripts`.
Make sure `output` is listed in your tsconfig's [`typeRoots`](https://www.typescriptlang.org/tsconfig#typeRoots) so the compiler sees them.

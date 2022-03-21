declare interface IEditExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'EditExtensionCommandSetStrings' {
  const strings: IEditExtensionCommandSetStrings;
  export = strings;
}

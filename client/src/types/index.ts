export interface Status {
    message: string;
    isError: boolean;
    isSuccess: boolean;
  }
  
  export interface FileData {
    name: string;
    file: File;
  }
// @ts-nocheck
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { SystemLists } from '../constants/SharePointListNames';

export interface IFileUploadProgress {
  fileName: string;
  status: "queued" | "processing" | "completed" | "failed";
  progress: number;
  error?: string;
  fileId?: number;
  fileUrl?: string;
}

export interface IBulkUploadResult {
  totalFiles: number;
  successCount: number;
  failureCount: number;
  files: IFileUploadProgress[];
  startTime: Date;
  endTime?: Date;
}

export interface IFileMetadata {
  fileName: string;
  fileContent: File | Blob;
  documentType?: string;
  policyCategory?: string;
  policyId?: number;
  metadata?: { [key: string]: any };
}

export class BulkFileUploadService {
  private sp: SPFI;
  private maxConcurrentUploads: number = 3;
  private uploadQueue: IFileMetadata[] = [];
  private activeUploads: number = 0;
  private results: IBulkUploadResult;
  private onProgressCallback?: (result: IBulkUploadResult) => void;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.results = {
      totalFiles: 0,
      successCount: 0,
      failureCount: 0,
      files: [],
      startTime: new Date()
    };
  }

  /**
   * Upload multiple files with progress tracking
   */
  public async uploadFiles(
    files: IFileMetadata[],
    onProgress?: (result: IBulkUploadResult) => void
  ): Promise<IBulkUploadResult> {
    this.onProgressCallback = onProgress;
    this.uploadQueue = [...files];

    this.results = {
      totalFiles: files.length,
      successCount: 0,
      failureCount: 0,
      files: files.map(f => ({
        fileName: f.fileName,
        status: "queued",
        progress: 0
      })),
      startTime: new Date()
    };

    // Notify initial state
    if (this.onProgressCallback) {
      this.onProgressCallback(this.results);
    }

    // Process queue with concurrency control
    await this.processQueue();

    this.results.endTime = new Date();

    // Final notification
    if (this.onProgressCallback) {
      this.onProgressCallback(this.results);
    }

    return this.results;
  }

  /**
   * Process the upload queue with concurrency control
   */
  private async processQueue(): Promise<void> {
    const promises: Promise<void>[] = [];

    while (this.uploadQueue.length > 0 || this.activeUploads > 0) {
      // Start new uploads up to the concurrency limit
      while (this.uploadQueue.length > 0 && this.activeUploads < this.maxConcurrentUploads) {
        const fileMetadata = this.uploadQueue.shift()!;
        this.activeUploads++;

        const promise = this.uploadSingleFile(fileMetadata)
          .finally(() => {
            this.activeUploads--;
          });

        promises.push(promise);
      }

      // Wait for at least one upload to complete before continuing
      if (promises.length > 0) {
        await Promise.race(promises);
      }
    }

    // Wait for all uploads to complete
    await Promise.all(promises);
  }

  /**
   * Upload a single file
   */
  private async uploadSingleFile(fileMetadata: IFileMetadata): Promise<void> {
    const fileProgress = this.results.files.find(f => f.fileName === fileMetadata.fileName);
    if (!fileProgress) return;

    try {
      // Update status to processing
      fileProgress.status = "processing";
      fileProgress.progress = 10;
      this.notifyProgress();

      // Detect file type if not provided
      const documentType = fileMetadata.documentType || this.detectFileType(fileMetadata.fileName);

      // Upload file to SharePoint
      fileProgress.progress = 30;
      this.notifyProgress();

      const uploadResult = await this.sp.web.lists
        .getByTitle("SystemLists.POLICY_SOURCE_DOCUMENTS")
        .rootFolder.files.addUsingPath(
          fileMetadata.fileName,
          fileMetadata.fileContent,
          { Overwrite: true }
        );

      fileProgress.progress = 60;
      this.notifyProgress();

      // Get the list item
      const item = await uploadResult.file.getItem();
      const itemId = (item as any).Id;

      fileProgress.progress = 70;
      this.notifyProgress();

      // Set metadata
      const metadata: any = {
        DocumentType: documentType,
        FileStatus: "Uploaded",
        UploadDate: new Date().toISOString(),
        PolicyCategory: fileMetadata.policyCategory || "General",
        ...(fileMetadata.metadata || {})
      };

      if (fileMetadata.policyId) {
        metadata.PolicyId = fileMetadata.policyId;
      }

      await item.update(metadata);

      fileProgress.progress = 90;
      this.notifyProgress();

      // Extract text content if applicable
      if (this.isTextExtractable(documentType)) {
        await this.queueForExtraction(itemId, fileMetadata.fileName, documentType);
      }

      // Get file URL
      const fileUrl = uploadResult.data.ServerRelativeUrl;

      // Mark as completed
      fileProgress.status = "completed";
      fileProgress.progress = 100;
      fileProgress.fileId = itemId;
      fileProgress.fileUrl = fileUrl;
      this.results.successCount++;
      this.notifyProgress();

    } catch (error) {
      console.error(`Failed to upload ${fileMetadata.fileName}:`, error);
      fileProgress.status = "failed";
      fileProgress.progress = 0;
      fileProgress.error = error instanceof Error ? error.message : "Upload failed";
      this.results.failureCount++;
      this.notifyProgress();
    }
  }

  /**
   * Queue file for text extraction
   */
  private async queueForExtraction(
    fileId: number,
    fileName: string,
    fileType: string
  ): Promise<void> {
    try {
      const queueData = {
        SourceFileUrl: {
          Url: `/sites/yoursite/SystemLists.POLICY_SOURCE_DOCUMENTS/${fileName}`,
          Description: fileName
        },
        SourceFileType: fileType,
        QueueStatus: "Queued",
        QueuedDate: new Date().toISOString()
      };

      await this.sp.web.lists
        .getByTitle("SystemLists.FILE_CONVERSION_QUEUE")
        .items.add(queueData);

    } catch (error) {
      console.error("Failed to queue for extraction:", error);
    }
  }

  /**
   * Detect file type from extension
   */
  private detectFileType(fileName: string): string {
    const extension = fileName.split(".").pop()?.toLowerCase() || "";

    const typeMap: { [key: string]: string } = {
      doc: "Word Document",
      docx: "Word Document",
      xls: "Excel Spreadsheet",
      xlsx: "Excel Spreadsheet",
      ppt: "PowerPoint Presentation",
      pptx: "PowerPoint Presentation",
      pdf: "PDF",
      jpg: "Image",
      jpeg: "Image",
      png: "Image",
      gif: "Image",
      bmp: "Image",
      mp4: "Video",
      avi: "Video",
      mov: "Video"
    };

    return typeMap[extension] || "Other";
  }

  /**
   * Check if file type supports text extraction
   */
  private isTextExtractable(fileType: string): boolean {
    const extractableTypes = [
      "Word Document",
      "Excel Spreadsheet",
      "PowerPoint Presentation",
      "PDF"
    ];
    return extractableTypes.includes(fileType);
  }

  /**
   * Notify progress callback
   */
  private notifyProgress(): void {
    if (this.onProgressCallback) {
      this.onProgressCallback({ ...this.results });
    }
  }

  /**
   * Cancel all pending uploads
   */
  public cancelUploads(): void {
    this.uploadQueue = [];

    // Mark all queued files as failed
    this.results.files
      .filter(f => f.status === "queued")
      .forEach(f => {
        f.status = "failed";
        f.error = "Upload cancelled by user";
        this.results.failureCount++;
      });

    this.notifyProgress();
  }

  /**
   * Get current upload status
   */
  public getUploadStatus(): IBulkUploadResult {
    return { ...this.results };
  }

  /**
   * Set maximum concurrent uploads
   */
  public setMaxConcurrentUploads(max: number): void {
    this.maxConcurrentUploads = Math.max(1, Math.min(max, 10)); // Between 1 and 10
  }

  /**
   * Validate file before upload
   */
  public static validateFile(file: File): { isValid: boolean; error?: string } {
    // Check file size (max 100MB)
    const maxSize = 100 * 1024 * 1024; // 100MB
    if (file.size > maxSize) {
      return {
        isValid: false,
        error: `File size exceeds maximum allowed size (100MB)`
      };
    }

    // Check file name
    if (!file.name || file.name.trim() === "") {
      return {
        isValid: false,
        error: "File name is required"
      };
    }

    // Check for invalid characters in file name
    const invalidChars = /[<>:"/\\|?*]/;
    if (invalidChars.test(file.name)) {
      return {
        isValid: false,
        error: "File name contains invalid characters"
      };
    }

    // Check file extension
    const allowedExtensions = [
      "doc", "docx", "xls", "xlsx", "ppt", "pptx",
      "pdf", "jpg", "jpeg", "png", "gif", "bmp",
      "mp4", "avi", "mov"
    ];

    const extension = file.name.split(".").pop()?.toLowerCase() || "";
    if (!allowedExtensions.includes(extension)) {
      return {
        isValid: false,
        error: `File type .${extension} is not allowed`
      };
    }

    return { isValid: true };
  }

  /**
   * Validate multiple files before upload
   */
  public static validateFiles(files: File[]): {
    validFiles: File[];
    invalidFiles: Array<{ file: File; error: string }>
  } {
    const validFiles: File[] = [];
    const invalidFiles: Array<{ file: File; error: string }> = [];

    for (const file of files) {
      const validation = this.validateFile(file);
      if (validation.isValid) {
        validFiles.push(file);
      } else {
        invalidFiles.push({ file, error: validation.error || "Invalid file" });
      }
    }

    return { validFiles, invalidFiles };
  }

  /**
   * Process conversion queue (to be called by a timer job or Power Automate)
   */
  public async processConversionQueue(): Promise<void> {
    try {
      const queuedItems = await this.sp.web.lists
        .getByTitle("SystemLists.FILE_CONVERSION_QUEUE")
        .items.filter("QueueStatus eq 'Queued'")
        .orderBy("QueuedDate", true)
        .top(10)();

      for (const item of queuedItems) {
        try {
          // Update status to processing
          await this.sp.web.lists
            .getByTitle("SystemLists.FILE_CONVERSION_QUEUE")
            .items.getById(item.Id)
            .update({ QueueStatus: "Processing" });

          // Here you would integrate with Azure Form Recognizer or similar service
          // For now, we'll just mark as completed with placeholder content
          const extractedContent = `Extracted content from ${item.SourceFileUrl.Description}`;

          await this.sp.web.lists
            .getByTitle("SystemLists.FILE_CONVERSION_QUEUE")
            .items.getById(item.Id)
            .update({
              QueueStatus: "Completed",
              ProcessedDate: new Date().toISOString(),
              ConvertedContent: extractedContent,
              ProcessingTime: Math.floor(Math.random() * 5000) + 1000 // Placeholder
            });

        } catch (error) {
          console.error(`Failed to process queue item ${item.Id}:`, error);
          await this.sp.web.lists
            .getByTitle("SystemLists.FILE_CONVERSION_QUEUE")
            .items.getById(item.Id)
            .update({
              QueueStatus: "Failed",
              ProcessedDate: new Date().toISOString(),
              ErrorMessage: error instanceof Error ? error.message : "Processing failed"
            });
        }
      }
    } catch (error) {
      console.error("Failed to process conversion queue:", error);
    }
  }
}

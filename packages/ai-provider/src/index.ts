export type {
  AiChatRequest,
  AiChatResponse,
  GenSparkAccountStatus,
  AiImageRequest,
  AiImageResponse,
  AiImageProviderConfig,
  AiImageProviderId,
  AiImageProviderMeta,
  AiImageProtocol,
  AiImageSettings,
  AiMediaProviderConfig,
  AiMediaProviderId,
  AiMediaProviderMeta,
  AiMediaSettings,
  AiAnalysisProtocol,
  AiMediaImageProtocol,
  AiSearchProviderId,
  AiSearchProviderMeta,
  AiSearchSettings,
  CodexModelCatalog,
  AiProviderConfig,
  AiModelCatalog,
  AiModelListRequest,
  AiModelListResponse,
  AiModelListErrorCode,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  LegacyAiSettings,
} from './types'
export type { IpcToolReply as AiToolReply } from '@genoffice/agent-core'
export {
  MODEL_CATALOG_TTL_MS,
  chatModelOptions,
  isModelId,
  normalizeModelIds,
  usableModelCatalog,
  modelCatalogBaseUrl,
} from './model-catalog'
export { listModelsForProvider } from './model-discovery'
export { registerAiModelDiscoveryIpc } from './model-discovery-ipc'
export {
  AI_PROVIDERS,
  DEFAULT_MAX_OUTPUT_TOKENS,
  DEFAULT_SYSTEM_PROMPT,
  MAX_MAX_OUTPUT_TOKENS,
  MIN_MAX_OUTPUT_TOKENS,
  clampMaxOutputTokens,
  cloudToolsEnabled,
  OPENAI_DEFAULT_BASE_URL,
  OPENAI_DEFAULT_IMAGE_MODEL,
  activeProvider,
  composeSystemPrompt,
  defaultAiSettings,
  maxOutputTokensOf,
  resolveAiSettings,
} from './providers'
export {
  AI_MEDIA_PROVIDERS,
  GEMINI_MEDIA_BASE_URL,
  OPENAI_IMAGES_BASE_URL,
  activeMediaConfig,
  activeMediaProvider,
  defaultAiMediaSettings,
  getMediaProviderMeta,
  imageGenerationAvailable,
  mediaAnalysisAvailable,
  mediaConfigUsable,
  providerHasCapability,
  resolveAiMediaSettings,
  videoAnalysisAvailable,
} from './media'
export type { MediaCapability } from './media'
export {
  AI_SEARCH_PROVIDERS,
  activeSearchProvider,
  defaultAiSearchSettings,
  resolveAiSearchSettings,
} from './search-settings'
export {
  analyzeMediaWithProvider,
  bytesToBase64,
  generateImageWithProvider,
  sniffImageMime,
  testMediaProvider,
} from './media-protocols'
export type {
  AnalyzeMediaInput,
  ByokMediaProviderId,
  GenerateImageInput,
  MediaBlob,
} from './media-protocols'
export {
  AI_PROVIDER_ADAPTERS,
  getProviderAdapter,
  modelLacksVision,
  endpointSupportsVision,
} from './registry'
export type {
  AiProtocol,
  ProviderAdapter,
  ProviderCapabilities,
  ResolvedEndpoint,
} from './registry'
export { chatForProvider } from './chat'
export { generateImageForProvider, generateImageForSettings } from './images'
export {
  AI_IMAGE_PROVIDERS,
  defaultImageSettings,
  imageConfigurationError,
  imageProviderMeta,
  resolveImageSettings,
} from './image-providers'
export type { AiImageOptions } from './images'
export { setAiUserAgent, setRescueFetch } from './fetch'
export { isAiNetworkError } from './network-error'
export { isAiOverloadedError } from './overload-error'
export { registerStandaloneAiIpc } from './ipc'
export type {
  AiIpcEventLike,
  AiAccountRoute,
  AiIpcHandler,
  AiIpcMainLike,
  AiIpcSenderLike,
  StandaloneAiIpcOptions,
} from './ipc'
export { parseOutputCapRejection } from './output-cap'
export { AiCreditsError, sseLines, streamForProvider } from './stream'
export type { StreamCallbacks } from './stream'
export {
  AI_ABSOLUTE_TIMEOUT_MS,
  AI_CHAT_RESPONSE_TIMEOUT_MS,
  AI_CONNECT_TIMEOUT_MS,
  AI_IDLE_TIMEOUT_MS,
  AiTimeoutError,
  createStreamWatchdog,
} from './watchdog'
export type { AiTimeoutKind, StreamWatchdog } from './watchdog'

export { isAiRegionBlockedError } from './region-error'

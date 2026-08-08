package com.oraldysplasia.ai.data.repository

import android.content.Context
import com.oraldysplasia.ai.data.local.AppDatabase
import com.oraldysplasia.ai.data.local.SlideEntity
import com.oraldysplasia.ai.data.remote.*
import com.oraldysplasia.ai.util.TokenManager
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.flow
import okhttp3.MediaType.Companion.toMediaTypeOrNull
import okhttp3.MultipartBody
import okhttp3.OkHttpClient
import okhttp3.RequestBody.Companion.asRequestBody
import okhttp3.RequestBody.Companion.toRequestBody
import okhttp3.logging.HttpLoggingInterceptor
import retrofit2.Retrofit
import retrofit2.converter.gson.GsonConverterFactory
import java.io.File
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch

class AppRepository(context: Context) {

    private val db = AppDatabase.getDatabase(context)
    private val slideDao = db.slideDao()
    val tokenManager = TokenManager(context)

    // Base URL uses 127.0.0.1 via ADB reverse port forwarding
    private val baseUrl = "http://127.0.0.1:8000/api/v1/"

    private val okHttpClient = OkHttpClient.Builder()
        .addInterceptor(AuthInterceptor(tokenManager))
        .addInterceptor(HttpLoggingInterceptor().setLevel(HttpLoggingInterceptor.Level.BODY))
        .build()

    private val apiService = Retrofit.Builder()
        .baseUrl(baseUrl)
        .client(okHttpClient)
        .addConverterFactory(GsonConverterFactory.create())
        .build()
        .create(ApiService::class.java)

    // ── Auth Methods ────────────────────────────────────────────────────
    suspend fun login(request: LoginRequest): Result<AuthResponse> {
        return try {
            val response = apiService.login(request)
            if (response.isSuccessful && response.body() != null) {
                val body = response.body()!!
                tokenManager.saveToken(body.access_token)
                tokenManager.saveUser(
                    body.user.id,
                    body.user.name,
                    body.user.email,
                    body.user.role,
                    body.user.institution,
                    body.user.license_id
                )
                Result.success(body)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Login failed"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    suspend fun signup(request: SignUpRequest): Result<AuthResponse> {
        return try {
            val response = apiService.signup(request)
            if (response.isSuccessful && response.body() != null) {
                val body = response.body()!!
                tokenManager.saveToken(body.access_token)
                tokenManager.saveUser(
                    body.user.id,
                    body.user.name,
                    body.user.email,
                    body.user.role,
                    body.user.institution,
                    body.user.license_id
                )
                Result.success(body)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Registration failed"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    suspend fun forgotPassword(email: String): Result<String> {
        return try {
            val response = apiService.forgotPassword(mapOf("email" to email))
            if (response.isSuccessful && response.body() != null) {
                Result.success(response.body()!!["message"] ?: "Recovery email sent")
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Email address not registered"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    fun logout() {
        tokenManager.clear()
        // Clear cached database as well to prevent cross-account leakage
        CoroutineScope(Dispatchers.IO).launch {
            try {
                slideDao.clearAll()
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }

    // ── Slides Methods (With caching logic) ─────────────────────────────
    fun getLibrary(
        page: Int = 1,
        limit: Int = 20,
        grade: String? = null,
        statusFilter: String? = null
    ): Flow<Result<List<SlideResponse>>> = flow {
        // Emit cached data first if any exists (only for page 1)
        if (page == 1) {
            val cache = slideDao.getAllSlides()
            if (cache.isNotEmpty()) {
                emit(Result.success(cache.map { it.toResponse() }))
            }
        }

        try {
            val response = apiService.getLibrary(page, limit, grade, statusFilter)
            if (response.isSuccessful && response.body() != null) {
                val networkList = response.body()!!.slides
                if (page == 1 && grade == null && statusFilter == null) {
                    slideDao.insertSlides(networkList.map { it.toEntity() })
                }
                emit(Result.success(networkList))
            } else {
                emit(Result.failure(Exception(response.errorBody()?.string() ?: "Network error")))
            }
        } catch (e: Exception) {
            emit(Result.failure(e))
        }
    }

    suspend fun uploadSlide(
        file: File,
        patientId: String,
        patientName: String,
        patientAge: String?,
        patientGender: String?,
        anatomicalSite: String,
        clinicalNotes: String?
    ): Result<SlideResponse> {
        return try {
            val filePart = MultipartBody.Part.createFormData(
                "file",
                file.name,
                file.asRequestBody("application/octet-stream".toMediaTypeOrNull())
            )
            val pId = patientId.toRequestBody("text/plain".toMediaTypeOrNull())
            val pName = patientName.toRequestBody("text/plain".toMediaTypeOrNull())
            val pAge = patientAge?.toRequestBody("text/plain".toMediaTypeOrNull())
            val pGen = patientGender?.toRequestBody("text/plain".toMediaTypeOrNull())
            val site = anatomicalSite.toRequestBody("text/plain".toMediaTypeOrNull())
            val notes = clinicalNotes?.toRequestBody("text/plain".toMediaTypeOrNull())

            val response = apiService.uploadSlide(filePart, pId, pName, pAge, pGen, site, notes)
            if (response.isSuccessful && response.body() != null) {
                val result = response.body()!!
                // Insert into cache db
                slideDao.insertSlide(result.toEntity())
                Result.success(result)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Upload failed"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    suspend fun getSlideDetail(slideId: Int): Result<SlideResponse> {
        // Try DB cache first
        val cached = slideDao.getSlideById(slideId)
        return try {
            val response = apiService.getSlideDetail(slideId)
            if (response.isSuccessful && response.body() != null) {
                val detail = response.body()!!
                slideDao.insertSlide(detail.toEntity())
                Result.success(detail)
            } else {
                cached?.let { Result.success(it.toResponse()) }
                    ?: Result.failure(Exception(response.errorBody()?.string() ?: "Failed to get slide details"))
            }
        } catch (e: Exception) {
            cached?.let { Result.success(it.toResponse()) } ?: Result.failure(e)
        }
    }

    suspend fun getDashboardStats(): Result<DashboardStats> {
        return try {
            val response = apiService.getDashboardStats()
            if (response.isSuccessful && response.body() != null) {
                Result.success(response.body()!!)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Failed to get dashboard statistics"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    // ── Analysis Methods ────────────────────────────────────────────────
    suspend fun runAnalysis(slideId: Int, confidenceThreshold: Float): Result<AnalysisRunResponse> {
        return try {
            val response = apiService.runAnalysis(AnalysisRunRequest(slide_id = slideId, confidence_threshold = confidenceThreshold))
            if (response.isSuccessful && response.body() != null) {
                Result.success(response.body()!!)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Failed to trigger analysis"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    suspend fun getAnalysisResult(slideId: Int): Result<AnalysisResultResponse> {
        return try {
            val response = apiService.getAnalysisResult(slideId)
            if (response.isSuccessful && response.body() != null) {
                Result.success(response.body()!!)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Failed to fetch analysis result"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    suspend fun submitReview(slideId: Int, finalGrade: String, icd10Code: String, comments: String?, annotations: List<Map<String, Any>>): Result<ReviewResponse> {
        return try {
            val response = apiService.submitReview(slideId, ReviewRequest(annotations, finalGrade, comments, icd10Code))
            if (response.isSuccessful && response.body() != null) {
                Result.success(response.body()!!)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Failed to submit review"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    // ── Reports Methods ─────────────────────────────────────────────────
    suspend fun exportReport(slideId: Int, format: String): Result<ReportResponse> {
        return try {
            val response = apiService.exportReport(slideId, format)
            if (response.isSuccessful && response.body() != null) {
                Result.success(response.body()!!)
            } else {
                Result.failure(Exception(response.errorBody()?.string() ?: "Failed to export report"))
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    // ── Mapping extension helper methods ────────────────────────────────
    private fun SlideResponse.toEntity() = SlideEntity(
        id = id,
        userId = user_id,
        patientId = patient_id,
        patientName = patient_name,
        patientAge = patient_age,
        patientGender = patient_gender,
        anatomicalSite = anatomical_site,
        filename = filename,
        sizeBytes = size_bytes,
        width = width,
        height = height,
        status = status,
        currentGrade = current_grade,
        overallConfidence = overall_confidence,
        clinicalNotes = clinical_notes,
        createdAt = created_at
    )

    private fun SlideEntity.toResponse() = SlideResponse(
        id = id,
        user_id = userId,
        patient_id = patientId,
        patient_name = patientName,
        patient_age = patientAge,
        patient_gender = patientGender,
        anatomical_site = anatomicalSite,
        filename = filename,
        size_bytes = sizeBytes,
        width = width,
        height = height,
        status = status,
        current_grade = currentGrade,
        overall_confidence = overallConfidence,
        clinical_notes = clinicalNotes,
        created_at = createdAt
    )
}

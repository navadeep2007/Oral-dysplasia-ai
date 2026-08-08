package com.oraldysplasia.ai.ui.screens.home

import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Add
import androidx.compose.material.icons.filled.AssignmentLate
import androidx.compose.material.icons.filled.FolderShared
import androidx.compose.material.icons.filled.HourglassEmpty
import androidx.compose.material.icons.filled.Search
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.lifecycle.viewmodel.compose.viewModel
import com.oraldysplasia.ai.data.repository.AppRepository
import com.oraldysplasia.ai.ui.components.GradeChip
import com.oraldysplasia.ai.ui.components.KpiCard
import com.oraldysplasia.ai.ui.theme.GradeSevere

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun HomeScreen(
    repository: AppRepository,
    onNavigateToUpload: () -> Unit,
    onNavigateToDetail: (Int) -> Unit
) {
    val viewModel: HomeViewModel = viewModel(
        factory = HomeViewModelFactory(repository)
    )

    LaunchedEffect(key1 = true) {
        viewModel.loadDashboard()
    }

    Scaffold(
        floatingActionButton = {
            FloatingActionButton(
                onClick = onNavigateToUpload,
                containerColor = Color(0xFF4F46E5), // Indigo FAB
                contentColor = Color.White,
                shape = RoundedCornerShape(16.dp)
            ) {
                Icon(Icons.Default.Add, contentDescription = "Upload Slide")
            }
        }
    ) { innerPadding ->
        Box(
            modifier = Modifier
                .fillMaxSize()
                .background(Color(0xFFF8FAFC)) // Soft Slate-50 background
                .padding(innerPadding)
        ) {
            when (val state = viewModel.state) {
                is HomeUiState.Loading -> {
                    Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                        CircularProgressIndicator(color = Color(0xFF4F46E5))
                    }
                }
                is HomeUiState.Error -> {
                    Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                        Column(horizontalAlignment = Alignment.CenterHorizontally, modifier = Modifier.padding(24.dp)) {
                            val isAuthError = state.message.contains("token", ignoreCase = true) ||
                                              state.message.contains("user not found", ignoreCase = true) ||
                                              state.message.contains("unauthorized", ignoreCase = true) ||
                                              state.message.contains("401", ignoreCase = true)
                            Text(
                                text = if (isAuthError) "Session expired or invalid credentials." else state.message,
                                color = MaterialTheme.colorScheme.error,
                                fontWeight = FontWeight.Bold
                            )
                            Spacer(modifier = Modifier.height(16.dp))
                            if (isAuthError) {
                                Button(
                                    onClick = {
                                        repository.tokenManager.clear()
                                        viewModel.loadDashboard()
                                    },
                                    colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF4F46E5))
                                ) {
                                    Text("Sign Out & Log In Again")
                                }
                            } else {
                                Button(onClick = { viewModel.loadDashboard() }) {
                                    Text("Retry")
                                }
                            }
                        }
                    }
                }
                is HomeUiState.Success -> {
                    val stats = state.stats
                    LazyColumn(
                        modifier = Modifier
                            .fillMaxSize()
                            .padding(horizontal = 16.dp),
                        verticalArrangement = Arrangement.spacedBy(16.dp)
                    ) {
                        // 1. Glowing Hero Banner Greeting
                        item {
                            Spacer(modifier = Modifier.height(16.dp))
                            Card(
                                modifier = Modifier.fillMaxWidth(),
                                shape = RoundedCornerShape(20.dp),
                                colors = CardDefaults.cardColors(containerColor = Color.Transparent)
                            ) {
                                Box(
                                    modifier = Modifier
                                        .fillMaxWidth()
                                        .background(
                                            Brush.horizontalGradient(
                                                colors = listOf(
                                                    Color(0xFF4F46E5), // Indigo-600
                                                    Color(0xFF0F766E)  // Teal-700
                                                )
                                            )
                                        )
                                        .padding(24.dp)
                                ) {
                                    Column {
                                        Text(
                                            text = "Welcome, Dr. ${viewModel.userName}",
                                            fontSize = 22.sp,
                                            fontWeight = FontWeight.ExtraBold,
                                            color = Color.White
                                        )
                                        Spacer(modifier = Modifier.height(4.dp))
                                        Text(
                                            text = viewModel.institution.uppercase(),
                                            fontSize = 11.sp,
                                            fontWeight = FontWeight.Bold,
                                            color = Color(0xFFCBD5E1),
                                            letterSpacing = 0.5.sp
                                        )
                                        Spacer(modifier = Modifier.height(12.dp))
                                        Text(
                                            text = "Oral Epithelial Dysplasia AI grading node active. Review diagnostic dossiers and verify histopathological findings below.",
                                            fontSize = 12.sp,
                                            color = Color(0xFFE2E8F0),
                                            lineHeight = 18.sp
                                        )
                                    }
                                }
                            }
                        }

                        // 2. Metrics KPIs Row
                        item {
                            Column(verticalArrangement = Arrangement.spacedBy(12.dp)) {
                                KpiCard(
                                    title = "Total Active Slides",
                                    value = stats.total_slides.toString(),
                                    icon = Icons.Default.FolderShared,
                                    iconColor = Color(0xFF4F46E5)
                                )
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.spacedBy(12.dp)
                                ) {
                                    KpiCard(
                                        title = "Pending Review",
                                        value = stats.pending_review.toString(),
                                        icon = Icons.Default.HourglassEmpty,
                                        iconColor = Color(0xFFD97706),
                                        modifier = Modifier.weight(1f)
                                    )
                                    KpiCard(
                                        title = "Severe Diagnoses",
                                        value = stats.severe_count.toString(),
                                        icon = Icons.Default.AssignmentLate,
                                        iconColor = GradeSevere,
                                        modifier = Modifier.weight(1f)
                                    )
                                }
                            }
                        }

                        // 3. Section Title
                        item {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Text(
                                    text = "Recent Biopsy Dossiers",
                                    fontSize = 17.sp,
                                    fontWeight = FontWeight.ExtraBold,
                                    color = Color(0xFF0F172A)
                                )
                                Icon(
                                    imageVector = Icons.Default.Search,
                                    contentDescription = null,
                                    tint = Color.Gray,
                                    modifier = Modifier.size(20.dp)
                                )
                            }
                        }

                        // 4. Biopsies List
                        if (stats.recent_slides.isEmpty()) {
                            item {
                                Card(
                                    modifier = Modifier.fillMaxWidth(),
                                    shape = RoundedCornerShape(12.dp),
                                    colors = CardDefaults.cardColors(containerColor = Color.White),
                                    border = BorderStroke(1.dp, Color(0xFFF1F5F9))
                                ) {
                                    Box(
                                        modifier = Modifier
                                            .fillMaxWidth()
                                            .padding(32.dp),
                                        contentAlignment = Alignment.Center
                                    ) {
                                        Text(
                                            text = "No biopsy cases uploaded to this node.",
                                            color = Color.Gray,
                                            fontSize = 13.sp,
                                            fontWeight = FontWeight.Medium
                                        )
                                    }
                                }
                            }
                        } else {
                            items(stats.recent_slides) { slide ->
                                Card(
                                    modifier = Modifier
                                        .fillMaxWidth()
                                        .clickable { onNavigateToDetail(slide.id) },
                                    shape = RoundedCornerShape(16.dp),
                                    colors = CardDefaults.cardColors(containerColor = Color.White),
                                    border = BorderStroke(1.dp, Color(0xFFE2E8F0))
                                ) {
                                    Row(
                                        modifier = Modifier
                                            .fillMaxWidth()
                                            .padding(16.dp),
                                        horizontalArrangement = Arrangement.SpaceBetween,
                                        verticalAlignment = Alignment.CenterVertically
                                    ) {
                                        Column(modifier = Modifier.weight(1f)) {
                                            Text(
                                                text = slide.filename,
                                                fontWeight = FontWeight.Bold,
                                                fontSize = 15.sp,
                                                color = Color(0xFF0F172A),
                                                maxLines = 1
                                            )
                                            Spacer(modifier = Modifier.height(6.dp))
                                            Row(verticalAlignment = Alignment.CenterVertically) {
                                                Text(
                                                    text = "Patient: ${slide.patient_id}",
                                                    fontSize = 11.sp,
                                                    fontWeight = FontWeight.SemiBold,
                                                    color = Color.Gray
                                                )
                                                Spacer(modifier = Modifier.width(12.dp))
                                                Text(
                                                    text = "Site: ${slide.anatomical_site}",
                                                    fontSize = 11.sp,
                                                    color = Color.Gray
                                                )
                                            }
                                        }
                                        Column(horizontalAlignment = Alignment.End) {
                                            GradeChip(grade = slide.current_grade)
                                            Spacer(modifier = Modifier.height(6.dp))
                                            val statusColor = when (slide.status) {
                                                "processed" -> Color(0xFF059669)
                                                "reviewed" -> Color(0xFF4F46E5)
                                                "analyzing" -> Color(0xFFD97706)
                                                else -> Color.Gray
                                            }
                                            Text(
                                                text = slide.status.uppercase(),
                                                fontSize = 9.sp,
                                                fontWeight = FontWeight.ExtraBold,
                                                color = statusColor,
                                                letterSpacing = 0.5.sp
                                            )
                                        }
                                    }
                                }
                            }
                        }

                        item {
                            Spacer(modifier = Modifier.height(40.dp))
                        }
                    }
                }
            }
        }
    }
}

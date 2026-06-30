# cat > /home/claude/expiry_project/devices/urls.py << 'EOF'
from django.urls import path
from . import views

urlpatterns = [
    path('update-expiry/', views.update_expiry_view, name='update_expiry'),
]
# EOF
# echo done
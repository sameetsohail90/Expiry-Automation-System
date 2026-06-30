import os
import tempfile
from django.shortcuts import render
from django.http import FileResponse

from .update_expiry import run_expiry_update


def update_expiry_view(request):
    context = {}

    if request.method == "POST" and request.FILES.get("excel_file"):
        uploaded_file = request.FILES["excel_file"]

        input_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(input_path, "wb+") as f:
            for chunk in uploaded_file.chunks():
                f.write(chunk)

        output_path = os.path.join(tempfile.gettempdir(), f"result_{uploaded_file.name}")

        # User ne agar date/time manually select ki ho to wo use karo, warna script
        # khud "+1 month from today" calculate kar legi.
        expiry_date = request.POST.get("expiry_date") or None   # e.g. "2026-07-30"
        expiry_time = request.POST.get("expiry_time") or "00:00"  # e.g. "14:30"

        manual_expiry = None
        if expiry_date:
            manual_expiry = f"{expiry_date} {expiry_time}:00"   # "2026-07-30 14:30:00"

        try:
            run_expiry_update(input_path, output_path, manual_expiry=manual_expiry)
        except Exception as e:
            context["error"] = str(e)
            return render(request, "expiry_app/update_expiry.html", context)

        return FileResponse(
            open(output_path, "rb"),
            as_attachment=True,
            filename=f"result_{uploaded_file.name}",
        )

    return render(request, "expiry_app/update_expiry.html", context)
# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from django.shortcuts import render
from django.http import HttpResponse
from django.views.decorators.csrf import csrf_exempt

from .forms import UploadFileForm


def index(request):
    context = {
        'latest_question_list': 'hello'
    }
    return render(request, 'excel/index.html', context)


@csrf_exempt
def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            mem_file = form.handleGroupExcel()
            resp = HttpResponse(mem_file, content_type='application/zip')
            resp['Content-Disposition'] = 'attachment; filename = output.zip'
            return resp
    else:
        form = UploadFileForm()
    return render(request, 'excel/index.html', {'form': form})
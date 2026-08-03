from io import BytesIO
from django.http import HttpResponse
from django.template.loader import get_template
from xhtml2pdf import pisa
from .models import (RepositoryIndicator,  StoreAPI)

def render_to_pdf(template_src, context_dict={}):
    template = get_template(template_src)
    html = template.render(context_dict)
    result = BytesIO()
    pdf = pisa.pisaDocument(BytesIO(html.encode("UTF-8")), result)
    if not pdf.err:
        return HttpResponse(result.getvalue(), content_type='application/pdf')
    return None




def bind_rows_data():
    store_data = StoreAPI.objects.all()
    repo_data = RepositoryIndicator.objects.all()

    unified_data = []

    # Normaliser StoreAPI
    for s in store_data:
        unified_data.append({
            "indicator_code": s.indicator_code,
            "country_code": s.country_code,
            "dim1_type": s.dim1_type,
            "dim1": s.dim1,
            "dim2_type": s.dim2_type,
            "dim2": s.dim2,
            "dim3_type": s.dim3_type,
            "dim3": s.dim3,
            "time_dim": s.time_dim,
            "alpha_value": s.alpha_value,
            "numeric_value": s.numeric_value,
            "publish_date": s.publish_date,
            "indicator_name": None,
        })

    # Normaliser RepositoryIndicator
    for r in repo_data:
        unified_data.append({
            "indicator_code": r.indicator_code,
            "country_code": r.spatial_dim,
            "dim1_type": r.dim1_type,
            "dim1": r.dim1,
            "dim2_type": r.dim2_type,
            "dim2": r.dim2,
            "dim3_type": r.dim3_type,
            "dim3": r.dim3,
            "time_dim": r.time_dim,
            "alpha_value": r.alpha_value,
            "numeric_value": float(r.numeric_value) if r.numeric_value else None,
            "publish_date": r.publish_date,
            "indicator_name": r.indicator.indicator_name if r.indicator else None,
        })

    return unified_data



# -*- coding: utf-8 -*-
import copy
import datetime
from decimal import Decimal
import re
from django.contrib.contenttypes import generic
from django.utils.formats import localize
from xlwt import Workbook, easyxf, XFStyle
from itertools import groupby

from django.http import HttpResponse
from django.shortcuts import render_to_response
from django.template import RequestContext
from django.utils.translation import ugettext_lazy as _
from django.db.models.fields import DateTimeField, DateField
from django.utils.encoding import force_unicode
from django.db.models import Q
from django import forms
from django.forms.models import fields_for_model
from django.db.models.related import RelatedObject
from django.conf import settings


from model_report.utils import base_label, ReportValue, ReportRow
from model_report.highcharts import HighchartRender
from model_report.widgets import RangeField
from model_report.export_pdf import render_to_pdf


import arial10


DEFAULT_CHART_TYPES = (
    ('area', _('Area')),
    ('line', _('Line')),
    ('column', _('Columns')),
    ('pie', _('Pie'))
)

CHART_SERIE_OPERATOR = (
    ('', u'---------'),
    ('sum', _('Sum')),
    ('len', _('Count')),
    ('avg', _('Average')),
    ('min', _('Min')),
    ('max', _('Max'))
)


class FitSheetWrapper(object):
    """Try to fit columns to max size of any entry.
    To use, wrap this around a worksheet returned from the
    workbook's add_sheet method, like follows:

        sheet = FitSheetWrapper(book.add_sheet(sheet_name))

    The worksheet interface remains the same: this is a drop-in wrapper
    for auto-sizing columns.
    """
    def __init__(self, sheet):
        self.sheet = sheet
        self.widths = dict()
        self.heights = dict()

    def write(self, r, c, label=u'', style=None):
        if not style:
            style = XFStyle()

        if isinstance(label, datetime.datetime):
            _saved_format = style.num_format_str
            style.num_format_str = 'dd/mm/yyyy hh:mm:ss'
            self.sheet.write(r, c, label, style)
            style.num_format_str = _saved_format
        elif isinstance(label, datetime.date):
            _saved_format = style.num_format_str
            style.num_format_str = 'dd/mm/yyyy'
            self.sheet.write(r, c, label, style)
            style.num_format_str = _saved_format
        # elif isinstance(label, (float, Decimal)):
        #     _saved_format = style.num_format_str
        #     style.num_format_str = '#,##0.00'
        #     self.sheet.write(r, c, label, style)
        #     style.num_format_str = _saved_format
        else:
            self.sheet.write(r, c, label, style)

        self.sheet.row(r).collapse = True

        unicode_label = unicode(label)

        bold = str(style.font.bold) in ('1', 'true', 'True')

        width = int(arial10.fitwidth(unicode_label, bold))
        if width > self.widths.get(c, 0):
            self.widths[c] = width
            self.sheet.col(c).width = width

        height = int(arial10.fitheight(unicode_label, bold))
        if height > self.heights.get(r, 0):
            self.heights[r] = height
            self.sheet.row(r).height = height

    def __getattr__(self, attr):
        return getattr(self.sheet, attr)


try:
    from collections import OrderedDict
except ImportError:
    OrderedDict = dict


# noinspection PyBroadException
def autodiscover():
    """
    Auto-discover INSTALLED_APPS report.py modules and fail silently when
    not present. Borrowed form django.contrib.admin
    """

    from django.utils.importlib import import_module
    from django.utils.module_loading import module_has_submodule

    global reports

    for app in settings.INSTALLED_APPS:
        mod = import_module(app)
        # Attempt to import the app's admin module.
        before_import_registry = copy.copy(reports)

        try:
            import_module('%s.reports' % app)
        except:
            # Reset the model registry to the state before the last import as
            # this import will have to reoccur on the next request and this
            # could raise NotRegistered and AlreadyRegistered exceptions
            # (see #8245).
            reports = before_import_registry

            # Decide whether to bubble up this error. If the app just
            # doesn't have an admin module, we can ignore the error
            # attempting to import it, otherwise we want it to bubble up.
            if module_has_submodule(mod, 'reports'):
                raise


class ReportClassManager(object):
    """
    Class to handle registered reports class.
    """

    _register = OrderedDict()

    def __init__(self):
        self._register = OrderedDict()

    def register(self, slug, rclass):
        if slug in self._register:
            raise ValueError('Slug already exists: %s' % slug)
        setattr(rclass, 'slug', slug)
        self._register[slug] = rclass

    def get_report(self, slug):
        # return class
        return self._register.get(slug, None)

    def get_reports(self):
        # return clasess
        return self._register.values()


reports = ReportClassManager()


_cache_class = {}


def cache_return(fun):
    """
    Usages of this decorator have been removed from the ReportAdmin base class.

    Caching method returns gets in the way of customization at the implementation level
    now that report instances can be modified based on request data.
    """
    def wrap(self, *args, **kwargs):
        cache_field = '%s_%s' % (self.__class__.__name__, fun.func_name)
        if cache_field in _cache_class:
            return _cache_class[cache_field]
        result = fun(self, *args, **kwargs)
        _cache_class[cache_field] = result
        return result
    return wrap


# noinspection PyProtectedMember
class ReportAdmin(object):
    """
    Class to represent a Report.
    """

    fields = []
    """List of fields or lookup fields for query results to be listed."""

    model = None
    """Primary django model to query."""

    list_filter = ()
    """List of fields or lookup fields to filter data."""

    list_order_by = ()
    """List of fields or lookup fields to order data."""

    list_group_by = ()
    """List of fields or lookup fields to group data."""

    list_serie_fields = ()
    """List of fields to group by results in chart."""

    template_name = 'model_report/report.html'
    """Template file name to render the report."""

    title = None
    """Title of the report."""

    type = 'report'
    """"report" for only report and  "chart" for report and chart graphic results."""

    group_totals = {}
    """Dictionary with field name as key and function to calculate their values."""

    report_totals = {}
    """Dictionary with field name as key and function to calculate their values."""

    override_field_values = {}
    """
    Dictionary with field name as key and function to parse their original values.

    ::

        override_field_values = {
            'men': men_format,
            'women': women_format
        }
    """
    override_field_formats = {}
    """Dictionary with field name as key and function to parse their value after :func:`override_field_values`."""

    override_field_labels = {}
    """Dictionary with field name as key and function to parse the column label."""

    override_field_choices = {}
    """#TODO"""

    override_field_filter_values = {}
    """#TODO"""

    override_group_value = {}
    """#TODO"""

    chart_types = ()
    """List of highchart types."""

    exports = ('excel', 'pdf')
    """Alternative render report as "pdf" or "csv"."""

    inlines = []
    """List of other's Report related to the main report."""

    queryset = None
    """#TODO"""

    onlytotals = False
    groupby = None
    slug = None

    def __init__(self, parent_report=None, request=None):
        self.parent_report = parent_report
        self.request = request
        model_fields = []
        model_m2m_fields = []
        self.related_inline_field = None
        self.related_inline_accessor = None
        self.related_fields = []
        for field in self.get_query_field_names():
            try:
                m2mfields = []
                if '__' in field:  # IF field has lookup
                    pre_field = None
                    base_model = self.model
                    for field_lookup in field.split("__"):
                        if not pre_field:
                            pre_field = base_model._meta.get_field_by_name(field_lookup)[0]
                            if 'ManyToManyField' in unicode(pre_field) or isinstance(pre_field, RelatedObject):
                                m2mfields.append(pre_field)
                        elif isinstance(pre_field, RelatedObject):
                            if isinstance(pre_field.field, generic.GenericRelation):
                                base_model = pre_field.parent_model
                            else:
                                base_model = pre_field.model
                            pre_field = base_model._meta.get_field_by_name(field_lookup)[0]
                        else:
                            if 'Date' in unicode(pre_field):
                                pre_field = pre_field
                            else:
                                base_model = pre_field.rel.to
                                pre_field = base_model._meta.get_field_by_name(field_lookup)[0]
                    model_field = pre_field
                else:
                    if not 'self.' in field:
                        model_field = self.model._meta.get_field_by_name(field)[0]
                    else:
                        get_attr = lambda s: getattr(s, field.split(".")[1])
                        get_attr.verbose_name = field
                        model_field = field
            except IndexError:
                raise ValueError('The field "%s" does not exist in model "%s".' % (field, self.model._meta.module_name))
            model_fields.append([model_field, field])
            if m2mfields:
                model_m2m_fields.append([model_field, field, len(model_fields) - 1, m2mfields])
        self.model_fields = model_fields
        self.model_m2m_fields = model_m2m_fields
        if parent_report:
            self.related_inline_field = [f for f, x in self.model._meta.get_fields_with_model()
                                         if f.rel and hasattr(f.rel, 'to') and f.rel.to is self.parent_report.model][0]
            self.related_inline_accessor = self.related_inline_field.related.get_accessor_name()
            self.related_fields = ["%s__%s" % (pfield.model._meta.module_name, attname) for pfield, attname in
                                   self.parent_report.model_fields if not isinstance(pfield, (str, unicode)) and
                                   pfield.model == self.related_inline_field.rel.to]
            self.related_inline_filters = []

            for pfield, pattname in self.parent_report.model_fields:
                for cfield, cattname in self.model_fields:
                    # try:
                        if pattname in cattname:
                            if pfield.model == cfield.model:
                                self.related_inline_filters.append([pattname, cattname,
                                                                    self.parent_report.get_fields().index(pattname)])
                    # TODO: narrow it
                    # except Exception, e:
                    #     pass

    def get_slug(self):
        if self.slug is None:
            self.slug = re.sub(r'(.)([A-Z])', r'\1-\2', self.__class__.__name__).lower()
        return self.slug

    def _get_grouper_text(self, groupby_field, value):
        # try:
        model_field = [mfield for mfield, field in self.model_fields if field == groupby_field][0]
        # TODO: narrow
        # except:
        #     model_field = None
        value = self.get_grouper_text(value, groupby_field, model_field)
        if value is None or unicode(value) == u'None':
            if groupby_field is None or unicode(groupby_field) == u'None':
                value = force_unicode(_('Results'))
            else:
                value = force_unicode(_('Nothing'))
        return value

    def _get_value_text(self, index, value, do_localize=True):
        # try:
        model_field = self.model_fields[index][0]
        # TODO: narrow
        # except:
        #     model_field = None

        value = self.get_value_text(value, index, model_field, do_localize)
        if value is None or unicode(value) == u'None':
            value = ''
        return value

    def get_grouper_text(self, value, field, model_field):
        if not isinstance(model_field, (str, unicode)) and '__' not in field:
            obj = model_field.model(**{field: value})
            if hasattr(obj, 'get_%s_display' % field):
                value = getattr(obj, 'get_%s_display' % field)()
        return value

    # @cache_return
    def get_m2m_field_names(self):
        return [field for ffield, field, index, mfield in self.model_m2m_fields]

    def get_value_text(self, value, index, model_field, do_localize=True):
        try:
            if not isinstance(model_field, (str, unicode)):
                obj = model_field.model(**{model_field.name: value})
                if hasattr(obj, 'get_%s_display' % model_field.name):
                    return getattr(obj, 'get_%s_display' % model_field.name)()
        except (TypeError, ValueError):
            pass
        return localize(value) if do_localize else value

    def get_empty_row_asdict(self, collection, default_value=None):
        erow = {}

        if default_value is None:
            default_value = []

        for field in collection:
            erow[field] = copy.copy(default_value)
        return dict(copy.deepcopy(erow))

    def reorder_dictrow(self, dictrow):
        return [dictrow[field_name] for field_name in self.get_fields()]

    def get_fields(self):
        return [x for x in self.fields if not x in self.related_fields]

    def get_column_names(self, ignore_columns=None):
        """
        Return the list of columns
        """
        values = []
        if ignore_columns is None:
            ignore_columns = {}
        for field, field_name in self.model_fields:
            if field_name in ignore_columns:
                continue
            caption = self.override_field_labels.get(field_name, base_label)(self, field)
            values.append(caption)
        return values

    # @cache_return
    def get_query_field_names(self):
        values = []
        for field in self.get_fields():
            if not 'self.' in field:
                values.append(field.split(".")[0])
            else:
                values.append(field)
        return values

    # @cache_return
    def get_queryset(self, filter_kwargs=None):
        """
        Return the the queryset
        """
        qs = self.model.objects.all()
        if filter_kwargs is not None:
            for k, v in filter_kwargs.items():
                if not v is None and v != '':
                    if hasattr(v, 'values_list'):
                        v = v.values_list('pk', flat=True)
                        k = '%s__pk__in' % k.split("__")[0]
                    qs = qs.filter(Q(**{k: v}))
        self.queryset = qs.distinct()
        return self.queryset

    def get_title(self):
        """
        Return the report title
        """
        title = self.title or None
        if not title:
            if not self.model:
                title = _('Unnamed report')
            else:
                title = force_unicode(self.model._meta.verbose_name_plural).lower().capitalize()
        return title

    def get_render_context(self, request, extra_context=None, by_row=None):
        context_request = request or self.request
        filter_related_fields = {}
        if self.parent_report and by_row:
            for mfield, cfield, index in self.related_inline_filters:
                filter_related_fields[cfield] = by_row[index].value

        try:
            form_groupby = self.get_form_groupby(context_request)
            form_filter = self.get_form_filter(context_request)
            form_config = self.get_form_config(context_request)

            column_labels = self.get_column_names(filter_related_fields)
            report_rows = []
            report_anchors = []
            chart = None

            # context = {
            #     'report': self,
            #     'form_groupby': form_groupby,
            #     'form_filter': form_filter,
            #     'form_config': form_config if self.type == 'chart' else None,
            #     'chart': chart,
            #     'report_anchors': report_anchors,
            #     'column_labels': column_labels,
            #     'report_rows': report_rows,
            # }

            if context_request.GET:
                groupby_data = form_groupby.get_cleaned_data() if form_groupby else {}
                filter_kwargs = filter_related_fields or form_filter.get_filter_kwargs()
                do_export = None
                do_localize = True

                if not context_request.GET.get('export', None) is None and not self.parent_report:
                    if context_request.GET.get('export') == 'excel':
                        do_export = 'excel'
                        do_localize = False
                    elif context_request.GET.get('export') == 'pdf':
                        do_export = 'pdf'

                if groupby_data:
                    # sets self.groupby and self.onlytotal variables
                    self.__dict__.update(groupby_data)

                report_rows = self.get_rows(groupby_data, filter_kwargs, filter_related_fields, do_localize=do_localize)

                for g, r in report_rows:
                    report_anchors.append(g)

                if len(report_anchors) <= 1:
                    report_anchors = []

                if self.type == 'chart' and groupby_data and 'groupby' in groupby_data:
                    config = form_config.get_config_data()
                    if config:
                        chart = self.get_chart(config, report_rows)

                if self.onlytotals:
                    for g, rows in report_rows:
                        for r in list(rows):
                            if r.is_value():
                                rows.remove(r)

                if do_export == 'excel':
                    book = Workbook(encoding='utf-8')
                    sheet1 = FitSheetWrapper(book.add_sheet(self.get_title()[:20]))
                    stylebold = easyxf('font: bold true; alignment:')
                    stylevalue = easyxf('alignment: horizontal left, vertical top;')
                    row_index = 0
                    for index, x in enumerate(column_labels):
                        sheet1.write(row_index, index, u'%s' % x, stylebold)
                    row_index += 1

                    for g, rows in report_rows:
                        if g:
                            sheet1.write(row_index, 0, unicode(g), stylebold)
                            row_index += 1
                        for row in list(rows):
                            if row.is_value():
                                for index, x in enumerate(row):
                                    # if isinstance(x.value, (list, tuple)):
                                    #     if len(x.value) < 1:
                                    #         xvalue = u''
                                    #     elif len(x.value) == 1:
                                    #         xvalue = x.value[0]
                                    #     else:
                                    #         xvalue = u''.join([unicode(v) for v in x.value])
                                    # else:
                                        # xvalue = x.text()
                                        # xvalue = x.value
                                    xvalue = x.formatted_value()
                                    sheet1.write(row_index, index, xvalue, stylevalue)
                                    # sheet1.write(row_index, index, x.value, stylevalue)
                                row_index += 1
                            elif row.is_caption:
                                for index, x in enumerate(row):
                                    if not isinstance(x, (unicode, str)):
                                        sheet1.write(row_index, index, x.text(), stylebold)
                                    else:
                                        sheet1.write(row_index, index, x, stylebold)
                                row_index += 1
                            elif row.is_total:
                                for index, x in enumerate(row):
                                    sheet1.write(row_index, index, x.text(), stylebold)
                                    sheet1.write(row_index + 1, index, u' ')
                                row_index += 2

                    response = HttpResponse(mimetype="application/ms-excel")
                    response['Content-Disposition'] = 'attachment; filename=%s.xls' % self.get_slug()
                    book.save(response)
                    return response

                if do_export == 'pdf':
                    inlines = [ir(self, context_request) for ir in self.inlines]
                    setattr(self, 'is_export', True)
                    context = {
                        'report': self,
                        'column_labels': column_labels,
                        'report_rows': report_rows,
                        'report_inlines': inlines,
                    }
                    context.update({'pagesize': 'legal landscape'})
                    return render_to_pdf(self, 'model_report/export_pdf.html', context)

            inlines = [ir(self, context_request) for ir in self.inlines]

            is_inline = self.parent_report is None
            render_report = not (len(report_rows) == 0 and is_inline)
            context = {
                'render_report': render_report,
                'is_inline': is_inline,
                'inline_column_span': 0 if is_inline else len(self.parent_report.get_column_names()),
                'report': self,
                'form_groupby': form_groupby,
                'form_filter': form_filter,
                'form_config': form_config if self.type == 'chart' else None,
                'chart': chart,
                'report_anchors': report_anchors,
                'column_labels': column_labels,
                'report_rows': report_rows,
                'report_inlines': inlines,
            }

            if extra_context:
                context.update(extra_context)

            context['request'] = request
            return context
        finally:
            globals()['_cache_class'] = {}

    def render(self, request, extra_context=None):
        context_or_response = self.get_render_context(request, extra_context)

        if isinstance(context_or_response, HttpResponse):
            return context_or_response
        return render_to_response(self.template_name, context_or_response, context_instance=RequestContext(request))

    def has_report_totals(self):
        return not (not self.report_totals)

    def has_group_totals(self):
        return not (not self.group_totals)

    def get_chart(self, config, report_rows):
        config['title'] = self.get_title()
        config['has_report_totals'] = self.has_report_totals()
        config['has_group_totals'] = self.has_group_totals()
        return HighchartRender(config).get_chart(report_rows)

    def get_form_config(self, request):

        class ConfigForm(forms.Form):

            chart_mode = forms.ChoiceField(label=_('Chart type'), choices=(), required=False)
            serie_field = forms.ChoiceField(label=_('Serie field'), choices=(), required=False)
            serie_op = forms.ChoiceField(label=_('Serie operator'), choices=CHART_SERIE_OPERATOR, required=False)

            def __init__(self, *args, **kwargs):
                super(ConfigForm, self).__init__(*args, **kwargs)
                choices = [('', '')]
                for k, v in DEFAULT_CHART_TYPES:
                    if k in self.chart_types:
                        choices.append((k, v))
                self.fields['chart_mode'].choices = list(choices)
                choices = [('', '')]
                for i, (index, mfield, field, caption) in enumerate(self.serie_fields):
                    choices += (
                        (index, caption),
                    )
                self.fields['serie_field'].choices = list(choices)

            def get_config_data(self):
                data = getattr(self, 'cleaned_data', {})
                if not data:
                    return {}
                if not data['serie_field'] or not data['chart_mode'] or not data['serie_op']:
                    return {}
                data['serie_field'] = int(data['serie_field'])
                return data

        ConfigForm.serie_fields = self.get_serie_fields()
        ConfigForm.chart_types = self.chart_types
        # ConfigForm.serie_fields
        form = ConfigForm(data=request.GET or None)
        form.is_valid()

        return form

    # @cache_return
    def get_groupby_fields(self):
        return [(mfield, field, caption) for (mfield, field), caption in zip(self.model_fields, self.get_column_names())
                if field in self.list_group_by]

    # @cache_return
    def get_serie_fields(self):
        return [(index, mfield, field, caption) for index, ((mfield, field), caption) in
                enumerate(zip(self.model_fields, self.get_column_names())) if field in self.list_serie_fields]

    # @cache_return
    def get_form_groupby(self, request):
        groupby_fields = self.get_groupby_fields()

        if not groupby_fields:
            return None

        class GroupByForm(forms.Form):

            groupby = forms.ChoiceField(label=_('Group by field:'), required=False)
            onlytotals = forms.BooleanField(label=_('Show only totals'), required=False)

            def _post_clean(self):
                pass

            def __init__(self, **kwargs):
                super(GroupByForm, self).__init__(**kwargs)
                choices = [(None, '')]
                for i, (mfield, field, caption) in enumerate(self.groupby_fields):
                    choices.append((field, caption))
                self.fields['groupby'].choices = choices
                data = kwargs.get('data', {})
                if data:
                    self.fields['groupby'].initial = data.get('groupby', '')

            def get_cleaned_data(self):
                cleaned_data = getattr(self, 'cleaned_data', {})
                if 'groupby' in cleaned_data:
                    if unicode(cleaned_data['groupby']) == u'None':
                        cleaned_data['groupby'] = None
                return cleaned_data

        GroupByForm.groupby_fields = groupby_fields

        form = GroupByForm(data=request.GET or None)
        form.is_valid()

        return form

    def get_form_filter(self, request):
        form_fields = fields_for_model(self.model, [f for f in self.get_query_field_names() if f in self.list_filter])
        if not form_fields:
            form_fields = {
                '__all__': forms.BooleanField(label='', widget=forms.HiddenInput, initial='1')
            }
        else:
            opts = self.model._meta
            for k, v in dict(form_fields).items():
                if v is None:
                    pre_field = None
                    base_model = self.model
                    if '__' in k:
                        for field_lookup in k.split("__")[:-1]:
                            if pre_field:
                                if isinstance(pre_field, RelatedObject):
                                    base_model = pre_field.model
                                else:
                                    base_model = pre_field.rel.to
                            pre_field = base_model._meta.get_field_by_name(field_lookup)[0]

                        model_field = pre_field
                    else:
                        field_name = k.split("__")[0]
                        model_field = opts.get_field_by_name(field_name)[0]

                    if isinstance(model_field, (DateField, DateTimeField)):
                        form_fields.pop(k)
                        field = RangeField(model_field.formfield)
                    else:
                        if not hasattr(model_field, 'formfield'):
                            field = forms.ModelChoiceField(queryset=model_field.model.objects.all())
                            if k in self.override_field_labels:
                                field.label = self.override_field_labels.get(k, base_label)(self, field)
                            # TODO find what is field lookup
                            # field.label = self.override_field_labels.get(k, base_label)(self, field) \
                            #     if k in self.override_field_labels else field_lookup
                        else:
                            field = model_field.formfield()
                        field.label = force_unicode(_(field.label))

                else:
                    if isinstance(v, forms.BooleanField):
                        form_fields.pop(k)
                        field = forms.ChoiceField()
                        field.label = v.label
                        field.help_text = v.help_text
                        field.choices = (
                            ('', ''),
                            (True, _('Yes')),
                            (False, _('No')),
                        )
                        setattr(field, 'as_boolean', True)
                    elif isinstance(v, (forms.DateField, forms.DateTimeField)):
                        field_name = k.split("__")[0]
                        model_field = opts.get_field_by_name(field_name)[0]
                        form_fields.pop(k)
                        field = RangeField(model_field.formfield)
                    else:
                        field = v

                    if hasattr(field, 'choices'):
                        # self.override_field_filter_values
                        if not hasattr(field, 'queryset'):
                            if field.choices[0][0]:
                                field.choices.insert(0, ('', '---------'))
                                field.initial = ''

                # Provide a hook for updating the queryset
                if hasattr(field, 'queryset') and k in self.override_field_choices:
                    field.queryset = self.override_field_choices.get(k)(self, field.queryset)
                form_fields[k] = field

        form_class = type('FilterFormBase', (forms.BaseForm,), {'base_fields': form_fields})

        # noinspection PyProtectedMember
        class FilterForm(form_class):

            def _post_clean(self):
                pass

            def get_filter_kwargs(self):
                if not self.is_valid():
                    return {}
                filter_kwargs = self.cleaned_data
                for key, value in filter_kwargs.items():
                    if not value:
                        filter_kwargs.pop(key)
                        continue
                    if key == '__all__':
                        filter_kwargs.pop(key)
                        continue
                    if isinstance(value, (list, tuple)):
                        if isinstance(self.fields[key], RangeField):
                            filter_kwargs.pop(key)
                            start_range, end_range = value
                            if start_range:
                                filter_kwargs['%s__gte' % key] = start_range
                            if end_range:
                                filter_kwargs['%s__lte' % key] = end_range
                    elif hasattr(self.fields[key], 'as_boolean'):
                        if value:
                            filter_kwargs.pop(key)
                            filter_kwargs[key] = (unicode(value) == u'True')
                return filter_kwargs

            def get_cleaned_data(self):
                return getattr(self, 'cleaned_data', {})

            def __init__(self, *args, **kwargs):
                super(FilterForm, self).__init__(*args, **kwargs)
                self.filter_report_is_all = '__all__' in self.fields and len(self.fields) == 1
                # try:
                data_filters = {}
                vals = args[0]
                for key in vals.keys():
                    if key in self.fields:
                        data_filters[key] = vals[key]
                for name in self.fields:
                    for key, value in data_filters.items():
                        if key == name:
                            continue
                        local_field = self.fields[name]
                        if hasattr(local_field, 'queryset'):
                            qs = local_field.queryset
                            if key in qs.model._meta.get_all_field_names():
                                local_field.queryset = qs.filter(Q(**{key: value}))
                # TODO narrow or remove
                # except:
                #     pass

                for local_field in self.fields:
                    self.fields[local_field].required = False

        form = FilterForm(data=request.GET or None)
        form.is_valid()

        return form

    def filter_query(self, qs):
        return qs

    def get_rows(self, groupby_data=None, filter_kwargs=None, filter_related_fields=None, do_localize=True):
        report_rows = []

        def get_field_value(obj, pfield):
            if isinstance(obj, dict):
                return obj[pfield]
            left_field = pfield.split("__")[0]
            # try:
            right_field = "__".join(pfield.split("__")[1:])
            # except:
            #     right_field = ''
            if right_field:
                return get_field_value(getattr(obj, left_field), right_field)
            if hasattr(obj, 'get_%s_display' % left_field):
                attr = getattr(obj, 'get_%s_display' % pfield)
            else:
                attr = getattr(obj, pfield)
            if callable(attr):
                attr = attr()
            return attr

        if filter_related_fields is None:
            filter_related_fields = {}

        if filter_kwargs is None:
            filter_kwargs = {}

        for kwarg, value in filter_kwargs.items():
            if kwarg in self.override_field_filter_values:
                filter_kwargs[kwarg] = self.override_field_filter_values.get(kwarg)(self, value)

        qs = self.get_queryset(filter_kwargs)
        ffields = ['pk' if f.startswith('self.') else f for f in self.get_query_field_names()
                   if f not in filter_related_fields]
        extra_ffield = []
        backend = settings.DATABASES['default']['ENGINE'].split('.')[-1]
        for f in list(ffields):
            if '__' in f:
                for field, name in self.model_fields:
                    if name == f:
                        if 'fields.Date' in unicode(field):
                            fname, flookup = f.rsplit('__', 1)
                            fname = fname.split('__')[-1]
                            if not flookup in ('year', 'month', 'day'):
                                break
                            if flookup == 'year':
                                if 'sqlite' in backend:
                                    extra_ffield.append([f, "strftime('%%Y', " + fname + ")"])
                                elif 'postgres' in backend:
                                    extra_ffield.append([f, "cast(extract(year from " + fname + ") as integer)"])
                                elif 'mysql' in backend:
                                    extra_ffield.append([f, "YEAR(" + fname + ")"])
                                else:
                                    raise NotImplemented  # mysql
                            if flookup == 'month':
                                if 'sqlite' in backend:
                                    extra_ffield.append([f, "strftime('%%m', " + fname + ")"])
                                elif 'postgres' in backend:
                                    extra_ffield.append([f, "cast(extract(month from " + fname + ") as integer)"])
                                elif 'mysql' in backend:
                                    extra_ffield.append([f, "MONTH(" + fname + ")"])
                                else:
                                    raise NotImplemented  # mysql
                            if flookup == 'day':
                                if 'sqlite' in backend:
                                    extra_ffield.append([f, "strftime('%%d', " + fname + ")"])
                                elif 'postgres' in backend:
                                    extra_ffield.append([f, "cast(extract(day from " + fname + ") as integer)"])
                                elif 'mysql' in backend:
                                    extra_ffield.append([f, "DAY(" + fname + ")"])
                                else:
                                    raise NotImplemented  # mysql
                        break
        obfields = list(self.list_order_by)
        if groupby_data and groupby_data['groupby']:
            if groupby_data['groupby'] in obfields:
                obfields.remove(groupby_data['groupby'])
            obfields.insert(0, groupby_data['groupby'])
        qs = self.filter_query(qs)
        qs = qs.order_by(*obfields)
        if extra_ffield:
            qs = qs.extra(select=dict(extra_ffield))
        qs = qs.values_list(*ffields)
        qs_list = list(qs)

        def get_with_dotvalues(resources):
            # {1: 'field.method'}
            dot_indexes = dict([(pos, dot_field) for pos, dot_field in enumerate(self.get_fields())
                                if '.' in dot_field])
            dot_indexes_values = {}

            dot_model_fields = [(pos, model_field[0]) for pos, model_field in enumerate(self.model_fields)
                                if pos in dot_indexes]
            # [ 1, model_field] ]
            for pos, model_field in dot_model_fields:
                model_ids = set([res[pos] for res in resources])
                if isinstance(model_field, (unicode, str)) and model_field.startswith('self.'):
                    model_qs = self.model.objects.filter(pk__in=model_ids)
                else:
                    model_qs = model_field.rel.to.objects.filter(pk__in=model_ids)
                div = {}
                method_name = dot_indexes[pos].split('.')[1]
                for obj in model_qs:
                    method_value = getattr(obj, method_name)
                    if callable(method_value):
                        method_value = method_value()
                    div[obj.pk] = method_value
                dot_indexes_values[pos] = div
                del model_qs

            if dot_indexes_values:
                new_resources = []
                for index_row, old_row in enumerate(resources):
                    new_row = []
                    for pos, actual_value in enumerate(old_row):
                        if pos in dot_indexes_values:
                            new_value = dot_indexes_values[pos][actual_value]
                        else:
                            new_value = actual_value
                        new_row.append(new_value)
                    new_resources.append(new_row)
                resources = new_resources
            return resources

        def compute_row_totals(row_config, row_values, is_group_total=False, is_report_total=False):
            total_row = self.get_empty_row_asdict(self.get_fields(), ReportValue(' '))
            for field_name in total_row.keys():
                if field_name in row_config:
                    fun = row_config[field_name]
                    cell_value = fun(row_values[field_name])
                    if field_name in self.get_m2m_field_names():
                        cell_value = ReportValue([cell_value])
                        # cell_value = [cell_value]
                    cell_value = ReportValue(cell_value)
                    cell_value.is_value = False
                    cell_value.is_group_total = is_group_total
                    cell_value.is_report_total = is_report_total
                    if field_name in self.override_field_values:
                        cell_value.to_value = self.override_field_values[field_name]
                    if field_name in self.override_field_formats:
                        cell_value.format = self.override_field_formats[field_name]
                    cell_value.is_m2m_value = field_name in self.get_m2m_field_names()
                    total_row[field_name] = cell_value
            totals_row = ReportRow(self.reorder_dictrow(total_row))
            totals_row.is_total = True
            return totals_row

        def compute_row_header(row_config):
            header_row = self.get_empty_row_asdict(self.get_fields(), ReportValue(''))
            for field_name, fun in row_config.items():
                if hasattr(fun, 'caption'):
                    field_value = force_unicode(fun.caption)
                else:
                    field_value = '&nbsp;'
                header_row[field_name] = field_value
            header_row = self.reorder_dictrow(header_row)
            header_row = ReportRow(header_row)
            header_row.is_caption = True
            return header_row

        def group_m2m_field_values(gqs_values):
            values_results = []
            m2m_indexes = [tpl[2] for tpl in self.model_m2m_fields]

            def get_key_values(gqs_vals):
                return [gv if field_idx not in m2m_indexes else None for field_idx, gv in enumerate(gqs_vals)]

            # gqs_values needs to already be sorted on the same key function
            # for groupby to work properly
            gqs_values.sort(key=get_key_values)
            res = groupby(gqs_values, key=get_key_values)

            for key, values in res:
                row_values = dict([(pos, []) for pos in m2m_indexes])
                for subrow_value in values:
                    for pos in m2m_indexes:
                        if subrow_value[pos] not in row_values[pos]:
                            row_values[pos].append(subrow_value[pos])
                for pos, vals in row_values.items():
                    key[pos] = vals
                values_results.append(key)
            return values_results

        qs_list = get_with_dotvalues(qs_list)
        if self.model_m2m_fields:
            qs_list = group_m2m_field_values(qs_list)

        if groupby_data and groupby_data['groupby']:
            groupby_field = groupby_data['groupby']
            if groupby_field in self.override_group_value:
                transform_fn = self.override_group_value.get(groupby_field)
                groupby_fn = lambda x: transform_fn(x[ffields.index(groupby_field)])
            else:
                groupby_fn = lambda x: x[ffields.index(groupby_field)]
        else:
            groupby_fn = lambda x: None

        qs_list.sort(key=groupby_fn)
        g = groupby(qs_list, key=groupby_fn)

        row_report_totals = self.get_empty_row_asdict(self.report_totals, [])
        for grouper, group_resources in g:
            rows = list()
            row_group_totals = self.get_empty_row_asdict(self.group_totals, [])
            for resource in group_resources:
                row = ReportRow()
                if isinstance(resource, (tuple, list)):
                    for index, value in enumerate(resource):
                        if ffields[index] in self.group_totals:
                            row_group_totals[ffields[index]].append(value)
                        elif ffields[index] in self.report_totals:
                            row_report_totals[ffields[index]].append(value)
                        value = self._get_value_text(index, value, do_localize=do_localize)
                        value = ReportValue(value)
                        if ffields[index] in self.override_field_values:
                            value.to_value = self.override_field_values[ffields[index]]
                        if ffields[index] in self.override_field_formats:
                            value.format = self.override_field_formats[ffields[index]]
                        row.append(value)
                else:
                    for index, column in enumerate(ffields):
                        value = get_field_value(resource, column)
                        if ffields[index] in self.group_totals:
                            row_group_totals[ffields[index]].append(value)
                        elif ffields[index] in self.report_totals:
                            row_report_totals[ffields[index]].append(value)
                        value = self._get_value_text(index, value, do_localize=do_localize)
                        value = ReportValue(value)
                        if column in self.override_field_values:
                            value.to_value = self.override_field_values[column]
                        if column in self.override_field_formats:
                            value.format = self.override_field_formats[column]
                        row.append(value)
                rows.append(row)
            if row_group_totals:
                if groupby_data['groupby']:
                    # header_group_total = compute_row_header(self.group_totals)
                    row = compute_row_totals(self.group_totals, row_group_totals, is_group_total=True)
                    # rows.append(header_group_total)
                    rows.append(row)
                for k, v in row_group_totals.items():
                    if k in row_report_totals:
                        row_report_totals[k].extend(v)

            if groupby_data and groupby_data['groupby']:
                grouper = self._get_grouper_text(groupby_data['groupby'], grouper)
            else:
                grouper = None
            if isinstance(grouper, (list, tuple)):
                grouper = grouper[0]
            report_rows.append([grouper, rows])
        if self.has_report_totals():
            header_report_total = compute_row_header(self.report_totals)
            row = compute_row_totals(self.report_totals, row_report_totals, is_report_total=True)
            header_report_total.is_report_totals = True
            row.is_report_totals = True
            report_rows.append([_('Totals'), [header_report_total, row]])

        return report_rows

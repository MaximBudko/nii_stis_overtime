class OverWorkController < ApplicationController
  def generate_overtime_report

  end

  def do_generate_ov
    # Параметры
    user_ids = params[:user_ids] || []
    @global_user_ids = user_ids
    start_date_str = params[:start_date]
    end_date_str = params[:end_date]
    type_overwork = params[:option_select]

    @global_type_overwork = type_overwork.to_s
    # Проверка выбранных пользователей
    if user_ids.blank?
      flash[:error] = "Пользователи не выбраны"
      redirect_to action: :generate_overtime_report and return 
    end

    if start_date_str.blank?
      flash[:error] = "Укажите дату начала"
      redirect_to action: :generate_overtime_report and return 
    end

    if end_date_str.blank?
      flash[:error] = "Укажите дату окончания"
      redirect_to action: :generate_overtime_report and return 
    end 
    start_date = Date.parse(start_date_str)
    end_date = Date.parse(end_date_str)
    
    
    date_range = (start_date..end_date).to_a
    @date_range_overtime = date_range

    # Проверка активности "Overtime"
    overtime_activity = TimeEntryActivity.find_by(name: "Сверхурочные")
    unless overtime_activity
      flash[:error] = "Активность 'Overtime' не найдена"
      redirect_to action: :generate_overtime_report and return 
    end

    # Поиск трудозатрат
    entries = TimeEntry.includes(:issue)
                       .where(user_id: user_ids, activity_id: overtime_activity.id)
    entries = entries.where(spent_on: date_range) if date_range

    # Сбор уникальных задач
    @overtime_issues = entries.map(&:issue).compact.uniq
    @users = User.where(id: user_ids)

    if @overtime_issues.blank?
      flash[:error] = "У выбранных пользователей нет задач с деятельностью 'Сверхурочные' за выбранный период."
      redirect_to action: :generate_overtime_report and return
    end

    respond_to do |format|
      format.xlsx do
        response.headers['Content-Disposition'] = 'attachment; filename=Отчет.xlsx'
        render xlsx: 'do_generate_ov', template: 'over_work/do_generate_ov'
      end
    end
  end

  private

  def load_form_data
    @users = User.all
  end
end

require 'docx'
require 'tempfile'

class OverWorkController < ApplicationController
  def generate_overtime_report

  end

  def users_by_date
    start_date = params[:start_date]
    end_date = params[:end_date]
  
    if start_date.blank? || end_date.blank?
      render json: [] and return
    end
  
    date_range = Date.parse(start_date)..Date.parse(end_date)
  
    # Находим кастомное поле "Тип работ"
    custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')
    unless custom_field
      render json: [] and return
    end
  
    # Получаем все трудозатраты с нужным типом работ
    overtime_entries = TimeEntry
      .includes(:user)
      .joins(:custom_values)
      .where(spent_on: date_range)
      .where(custom_values: { custom_field_id: custom_field.id, value: 'Сверхурочная' })
  
    # Группируем по пользователям
    user_entries = overtime_entries.group_by(&:user)
  
    # Формируем JSON
    users_json = user_entries.map do |user, entries|
      {
        id: user.id,
        name: "#{user.firstname} #{user.lastname}",
        dates: entries.map(&:spent_on).uniq.sort.map(&:to_s)
      }
    end
  
    render json: users_json
  end
  
  

  def do_generate_ov
    user_ids = params[:user_ids] || []
    @global_user_ids = user_ids
    start_date_str = params[:start_date]
    end_date_str = params[:end_date]
    type_overwork = params[:option_select]

    @global_type_overwork = type_overwork.to_s

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

    custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')

    entries = TimeEntry
      .includes(:issue)
      .joins(:custom_values)
      .where(user_id: user_ids)
      .where(spent_on: date_range)
      .where(custom_values: {
        custom_field_id: custom_field.id,
        value: 'Сверхурочная'
      })

    @overtime_issues = entries.map(&:issue).compact.uniq
    @users = User.where(id: user_ids)

    if @overtime_issues.blank?
      flash[:error] = "У выбранных пользователей нет задач с типом работ 'Сверхурочная' за выбранный период."
      redirect_to action: :generate_overtime_report and return
    end

    respond_to do |format|
      format.xlsx do
        response.headers['Content-Disposition'] = 'attachment; filename=Отчет.xlsx'
        render xlsx: 'Отчет о сверхурочной работе', template: 'over_work/do_generate_ov'
      end
    end
  end


  # def do_dokladnaya
  #   user_ids = params[:user_ids] || []
  #   @global_user_ids = user_ids
  #   start_date_str = params[:start_date]
  #   end_date_str = params[:end_date]
  #   type_overwork_dokladnaya = params[:option_select]

  #   @global_type_overwork_dokladnaya = type_overwork_dokladnaya.to_s
    
  #   p @global_type_overwork_dokladnaya

  #   if user_ids.blank?
  #     flash[:error] = "Пользователи не выбраны"
  #     redirect_to action: :generate_overtime_report and return 
  #   end

  #   if start_date_str.blank?
  #     flash[:error] = "Укажите дату начала"
  #     redirect_to action: :generate_overtime_report and return 
  #   end

  #   if end_date_str.blank?
  #     flash[:error] = "Укажите дату окончания"
  #     redirect_to action: :generate_overtime_report and return 
  #   end 

  #   start_date = Date.parse(start_date_str)
  #   end_date = Date.parse(end_date_str)

  #   date_range = (start_date..end_date).to_a
  #   @date_range_overtime = date_range

  #   custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')

  #   entries = TimeEntry
  #     .includes(:issue)
  #     .joins(:custom_values)
  #     .where(user_id: user_ids)
  #     .where(spent_on: date_range)
  #     .where(custom_values: {
  #       custom_field_id: custom_field.id,
  #       value: 'Сверхурочная'
  #     })

  #   @overtime_issues = entries.map(&:issue).compact.uniq
  #   @users = User.where(id: user_ids)

  #   if @overtime_issues.blank?
  #     flash[:error] = "У выбранных пользователей нет задач с типом работ 'Сверхурочная' за выбранный период."
  #     redirect_to action: :generate_overtime_report and return
  #   end

  #   respond_to do |format|
  #     format.xlsx do
  #       response.headers['Content-Disposition'] = 'attachment; filename=Отчет.xlsx'
  #       render xlsx: 'do_dokladnaya', template: 'over_work/do_dokladnaya'
  #     end
  #   end
  # end
  
  private

  def load_form_data
    @users = User.all
  end
end
